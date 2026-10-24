"""Flatten a PowerPoint (.pptx family) deck so it renders as static slides.

A .pptx is a ZIP of XML parts. This module rewrites the slide parts so that
headless LibreOffice renders what a viewer would see:

* every ``<p:timing>`` block (entrance, exit, emphasis and motion-path
  animations, build lists) and every ``<p:transition>`` is removed from
  slides, layouts and masters;
* the ``hidden="1"`` flag is cleared on slide shapes so hidden and
  click-triggered objects are rendered;
* the main animation sequence is replayed to find which shapes and
  paragraphs are on screen. By default only the **final** state is kept:
  one slide per original slide, with shapes that have exited removed and
  paragraphs that have exited blanked. With ``split_steps=True`` the slide
  is instead duplicated into one slide per click: the state before the
  first click, then the state after each click.

Click-triggered (interactive) sequences are treated as "shown": the shapes
they target are kept and unhidden, since a reader of the PDF cannot click.

Everything else in the package is copied through byte for byte. lxml is used
rather than ``xml.etree`` because it preserves namespace prefixes and
``mc:Ignorable`` declarations, which PowerPoint and LibreOffice both rely on.
"""
from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P, A, R, REL, CT = (f"{{{ns}}}" for ns in (P_NS, A_NS, R_NS, REL_NS, CT_NS))

SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
NOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"

SLIDE_PART = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
TIMED_PART = re.compile(
    r"^ppt/(?:slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$"
)
CONTENT_TYPES = "[Content_Types].xml"
PRESENTATION = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"

# Elements in the shape tree that carry a <p:nvXxxPr><p:cNvPr id=.../> block.
SHAPE_TAGS = {P + t for t in ("sp", "grpSp", "pic", "graphicFrame", "cxnSp")}
# Effects with these node types start on their own, without a click.
AUTO_START = {"withEffect", "afterEffect"}

_PARSER = etree.XMLParser(resolve_entities=False, remove_blank_text=False, huge_tree=True)


class FlattenError(Exception):
    """The package could not be read or rewritten as a .pptx."""


@dataclass(frozen=True)
class State:
    """What is *not* on screen at one moment of a slide."""

    hidden_shapes: frozenset[str] = frozenset()
    hidden_paragraphs: frozenset[tuple[str, int]] = frozenset()


@dataclass
class FlattenStats:
    slides: int = 0
    slides_added: int = 0
    timings_removed: int = 0
    transitions_removed: int = 0
    shapes_unhidden: int = 0
    shapes_removed: int = 0
    paragraphs_blanked: int = 0

    @property
    def changed(self) -> bool:
        return bool(
            self.slides_added
            or self.timings_removed
            or self.transitions_removed
            or self.shapes_unhidden
            or self.shapes_removed
            or self.paragraphs_blanked
        )

    def summary(self) -> str:
        parts = []
        if self.timings_removed:
            parts.append(f"{self.timings_removed} animation block(s) stripped")
        if self.slides_added:
            parts.append(f"{self.slides_added} step slide(s) added")
        if self.shapes_unhidden:
            parts.append(f"{self.shapes_unhidden} hidden shape(s) revealed")
        if self.shapes_removed:
            parts.append(f"{self.shapes_removed} exited shape(s) removed")
        if self.paragraphs_blanked:
            parts.append(f"{self.paragraphs_blanked} exited paragraph(s) blanked")
        if self.transitions_removed:
            parts.append(f"{self.transitions_removed} transition(s) dropped")
        return ", ".join(parts) if parts else "nothing to flatten"


# --------------------------------------------------------------------------- #
# package level                                                                #
# --------------------------------------------------------------------------- #


def flatten_pptx(src: Path, dst: Path, *, split_steps: bool = False) -> FlattenStats:
    """Write a flattened copy of ``src`` to ``dst`` and report what changed."""
    stats = FlattenStats()
    try:
        with zipfile.ZipFile(src) as zin:
            infos = zin.infolist()
            parts = {info.filename: zin.read(info) for info in infos}
    except zipfile.BadZipFile as exc:
        raise FlattenError(f"not a valid .pptx package: {exc}") from exc

    try:
        new_parts = _flatten_package(parts, stats, split_steps)
    except etree.XMLSyntaxError as exc:
        raise FlattenError(f"malformed slide XML: {exc}") from exc
    except (KeyError, ValueError) as exc:
        raise FlattenError(f"unexpected package layout: {exc}") from exc

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, parts[info.filename])
        for name, data in new_parts:
            zout.writestr(zipfile.ZipInfo(name, date_time=infos[0].date_time), data, zipfile.ZIP_DEFLATED)
    return stats


def _flatten_package(parts: dict[str, bytes], stats: FlattenStats, split_steps: bool) -> list[tuple[str, bytes]]:
    """Rewrite ``parts`` in place; return extra (name, data) parts to append."""
    slide_numbers = [int(m.group(1)) for name in parts if (m := SLIDE_PART.match(name))]
    next_number = max(slide_numbers, default=0) + 1
    added: list[tuple[str, bytes]] = []
    extra_slides: list[tuple[str, list[str]]] = []  # (original part, [new parts])

    for name in list(parts):
        if not TIMED_PART.match(name):
            continue
        if not SLIDE_PART.match(name):
            parts[name] = flatten_part(parts[name], is_slide=False, stats=stats)
            continue
        renders = flatten_slide(parts[name], stats=stats, split_steps=split_steps)
        parts[name] = renders[0]
        if len(renders) == 1:
            continue
        rels_name = _rels_name(name)
        rels = parts.get(rels_name)
        new_names = []
        for render in renders[1:]:
            new_name = f"ppt/slides/slide{next_number}.xml"
            next_number += 1
            added.append((new_name, render))
            if rels is not None:
                added.append((_rels_name(new_name), _rels_without_notes(rels)))
            new_names.append(new_name)
        extra_slides.append((name, new_names))
        stats.slides_added += len(new_names)

    if extra_slides:
        _register_slides(parts, extra_slides)
    return added


def _rels_name(part: str) -> str:
    folder, base = posixpath.split(part)
    return f"{folder}/_rels/{base}.rels"


def _rels_without_notes(rels_xml: bytes) -> bytes:
    """A duplicated slide must not claim the original's notes page."""
    root = etree.fromstring(rels_xml, _PARSER)
    for rel in list(root):
        if rel.get("Type") == NOTES_REL_TYPE:
            root.remove(rel)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _register_slides(parts: dict[str, bytes], extra_slides: list[tuple[str, list[str]]]) -> None:
    """Make presentation.xml, its rels and [Content_Types].xml know the new slides."""
    types = etree.fromstring(parts[CONTENT_TYPES], _PARSER)
    rels = etree.fromstring(parts[PRESENTATION_RELS], _PARSER)
    pres = etree.fromstring(parts[PRESENTATION], _PARSER)
    sld_list = pres.find(P + "sldIdLst")
    if sld_list is None:
        raise ValueError("presentation.xml has no sldIdLst")

    rel_targets = {rel.get("Id"): posixpath.normpath(posixpath.join("ppt", rel.get("Target", "").lstrip("/")))
                   for rel in rels}
    rel_by_target = {target: rid for rid, target in rel_targets.items()}
    rid_numbers = [int(m.group(1)) for rid in rel_targets if (m := re.match(r"rId(\d+)$", rid))]
    next_rid = max(rid_numbers, default=0) + 1
    next_id = max((int(s.get("id", "0")) for s in sld_list), default=255) + 1

    for original, new_names in extra_slides:
        rid = rel_by_target.get(original)
        anchor = next((s for s in sld_list if s.get(R + "id") == rid), None)
        if anchor is None:
            raise ValueError(f"{original} is not listed in presentation.xml")
        for new_name in new_names:
            new_rid = f"rId{next_rid}"
            next_rid += 1
            rel = etree.SubElement(rels, REL + "Relationship")
            rel.set("Id", new_rid)
            rel.set("Type", SLIDE_REL_TYPE)
            rel.set("Target", posixpath.relpath(new_name, "ppt"))

            override = etree.SubElement(types, CT + "Override")
            override.set("PartName", "/" + new_name)
            override.set("ContentType", SLIDE_CONTENT_TYPE)

            sld = etree.Element(P + "sldId")
            sld.set("id", str(next_id))
            sld.set(R + "id", new_rid)
            next_id += 1
            anchor.addnext(sld)
            anchor = sld

    for name, root in ((CONTENT_TYPES, types), (PRESENTATION_RELS, rels), (PRESENTATION, pres)):
        parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------- #
# part level                                                                   #
# --------------------------------------------------------------------------- #


def flatten_part(xml: bytes, *, is_slide: bool, stats: FlattenStats | None = None) -> bytes:
    """Flatten one part to its final state and return the new XML bytes."""
    if is_slide:
        return flatten_slide(xml, stats=stats)[0]
    stats = stats if stats is not None else FlattenStats()
    root = etree.fromstring(xml, _PARSER)
    _strip_timing(root, stats)
    return _serialize(root)


def flatten_slide(xml: bytes, *, stats: FlattenStats | None = None, split_steps: bool = False) -> list[bytes]:
    """Flatten one slide; return one XML part per rendered state.

    Without ``split_steps`` the list has exactly one entry, the final state.
    """
    stats = stats if stats is not None else FlattenStats()
    stats.slides += 1
    root = etree.fromstring(xml, _PARSER)
    timing = _strip_timing(root, stats)
    states = build_steps(timing) if timing is not None else [State()]
    if not split_steps:
        states = states[-1:]

    spTree = root.find(f"{P}cSld/{P}spTree")
    if spTree is not None:
        stats.shapes_unhidden += _unhide_shapes(spTree)

    renders = []
    for state in states:
        page = copy.deepcopy(root) if len(states) > 1 else root
        tree = page.find(f"{P}cSld/{P}spTree")
        if tree is not None:
            stats.shapes_removed += _remove_shapes(tree, state.hidden_shapes)
            stats.paragraphs_blanked += _blank_paragraphs(tree, state.hidden_paragraphs)
        renders.append(_serialize(page))
    return renders


def _strip_timing(root: etree._Element, stats: FlattenStats) -> etree._Element | None:
    """Detach and return the slide's timing tree; drop its transition."""
    timing = root.find(P + "timing")
    if timing is not None:
        timing.getparent().remove(timing)
        stats.timings_removed += 1
    transition = root.find(P + "transition")
    if transition is not None:
        transition.getparent().remove(transition)
        stats.transitions_removed += 1
    return timing


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# --------------------------------------------------------------------------- #
# animation replay                                                             #
# --------------------------------------------------------------------------- #

Effect = tuple[str, str, tuple[int, ...] | None]  # presetClass, shape id, paragraphs (None = whole shape)
Key = tuple[str, int | None]


def build_steps(timing: etree._Element) -> list[State]:
    """Replay the main sequence: the state before the first click, then after each click.

    Top-level groups under the main sequence are clicks, in play order. A group
    whose first effect starts automatically (with/after previous) is folded
    into the state before it. Groups that change nothing visible (emphasis,
    motion paths) add no state. Paragraph-level effects track individual
    paragraphs; other sub-shape targets (chart series, backgrounds) are ignored.
    """
    main = next((c for c in timing.iter(P + "cTn") if c.get("nodeType") == "mainSeq"), None)
    if main is None:
        return [State()]
    children = main.find(P + "childTnLst")
    groups = list(children) if children is not None else []
    per_group = [_effects(group) for group in groups]

    visible: dict[Key, bool] = {}
    for effects in per_group:
        for preset, spid, paragraphs in effects:
            for key in _keys(spid, paragraphs):
                visible.setdefault(key, preset != "entr")  # entering shapes start hidden

    states = [_snapshot(visible)]
    for group, effects in zip(groups, per_group):
        if not effects:
            continue
        for preset, spid, paragraphs in effects:
            for key in _keys(spid, paragraphs):
                visible[key] = preset == "entr"
        snapshot = _snapshot(visible)
        if _starts_automatically(group):
            states[-1] = snapshot
        elif snapshot != states[-1]:
            states.append(snapshot)
    return states


def _effects(group: etree._Element) -> list[Effect]:
    found: list[Effect] = []
    for ctn in group.iter(P + "cTn"):
        preset = ctn.get("presetClass")
        if preset not in ("entr", "exit"):
            continue
        for behaviour in ctn.iter(P + "cBhvr"):
            target = behaviour.find(f"{P}tgtEl/{P}spTgt")
            if target is None or target.get("spid") is None:
                continue
            spid = target.get("spid")
            if len(target) == 0:
                found.append((preset, spid, None))
                continue
            paragraph_range = target.find(f"{P}txEl/{P}pRg")
            if paragraph_range is None:
                continue  # character ranges, chart series, backgrounds: not modelled
            start, end = int(paragraph_range.get("st")), int(paragraph_range.get("end"))
            found.append((preset, spid, tuple(range(start, end + 1))))
    return found


def _keys(spid: str, paragraphs: tuple[int, ...] | None) -> list[Key]:
    return [(spid, None)] if paragraphs is None else [(spid, i) for i in paragraphs]


def _snapshot(visible: dict[Key, bool]) -> State:
    return State(
        hidden_shapes=frozenset(spid for (spid, i), on in visible.items() if i is None and not on),
        hidden_paragraphs=frozenset((spid, i) for (spid, i), on in visible.items() if i is not None and not on),
    )


def _starts_automatically(group: etree._Element) -> bool:
    first = next((c for c in group.iter(P + "cTn") if c.get("presetClass")), None)
    return first is not None and first.get("nodeType") in AUTO_START


# --------------------------------------------------------------------------- #
# shape tree edits                                                             #
# --------------------------------------------------------------------------- #


def _shape_of(cNvPr: etree._Element) -> etree._Element | None:
    nv = cNvPr.getparent()
    shape = nv.getparent() if nv is not None else None
    return shape if shape is not None and shape.tag in SHAPE_TAGS else None


def _shapes_by_id(spTree: etree._Element, spids: set[str] | frozenset[str]) -> dict[str, etree._Element]:
    found = {}
    for cNvPr in spTree.iter(P + "cNvPr"):
        if cNvPr.get("id") in spids:
            shape = _shape_of(cNvPr)
            if shape is not None:
                found[cNvPr.get("id")] = shape
    return found


def _remove_shapes(spTree: etree._Element, spids: frozenset[str]) -> int:
    removed = 0
    for shape in _shapes_by_id(spTree, spids).values():
        if shape.getparent() is not None:
            shape.getparent().remove(shape)
            removed += 1
    return removed


def _blank_paragraphs(spTree: etree._Element, hidden: frozenset[tuple[str, int]]) -> int:
    """Empty the text of paragraphs that are off screen, keeping their spacing."""
    if not hidden:
        return 0
    blanked = 0
    by_shape: dict[str, set[int]] = {}
    for spid, index in hidden:
        by_shape.setdefault(spid, set()).add(index)
    for spid, shape in _shapes_by_id(spTree, set(by_shape)).items():
        body = shape.find(P + "txBody")
        if body is None:
            continue
        paragraphs = body.findall(A + "p")
        for index in by_shape[spid]:
            if index >= len(paragraphs):
                continue
            for child in list(paragraphs[index]):
                if child.tag not in (A + "pPr", A + "endParaRPr"):
                    paragraphs[index].remove(child)
            blanked += 1
    return blanked


def _unhide_shapes(spTree: etree._Element) -> int:
    unhidden = 0
    for cNvPr in spTree.iter(P + "cNvPr"):
        if cNvPr.get("hidden") is not None and _shape_of(cNvPr) is not None:
            del cNvPr.attrib["hidden"]
            unhidden += 1
    return unhidden
