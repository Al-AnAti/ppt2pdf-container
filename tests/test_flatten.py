"""Unit tests for flatten.py (no LibreOffice needed)."""
import zipfile
from pathlib import Path

import pytest
from lxml import etree

import flatten
from flatten import A, P, R, FlattenError, FlattenStats, State, build_steps, flatten_part, flatten_pptx

FIXTURE = Path(__file__).parent / "fixtures" / "animated.pptx"


def shapes_by_name(xml: bytes) -> dict[str, etree._Element]:
    root = etree.fromstring(xml)
    return {el.get("name"): el for el in root.iter(P + "cNvPr") if el.get("name")}


def texts(xml: bytes) -> list[str]:
    return [t.text or "" for t in etree.fromstring(xml).iter(A + "t")]


def read(pptx: Path, name: str) -> bytes:
    with zipfile.ZipFile(pptx) as z:
        return z.read(name)


# --------------------------------------------------------------------------- #
# default mode: final state, one slide per slide                               #
# --------------------------------------------------------------------------- #


@pytest.fixture(scope="module")
def flattened(tmp_path_factory):
    out = tmp_path_factory.mktemp("flat") / "animated.pptx"
    stats = flatten_pptx(FIXTURE, out)
    return out, stats


def test_stats_report_every_change(flattened):
    _, stats = flattened
    assert stats.slides == 2
    assert stats.slides_added == 0
    assert stats.timings_removed == 2
    assert stats.transitions_removed == 1
    assert stats.shapes_unhidden == 1
    assert stats.shapes_removed == 1
    assert stats.paragraphs_blanked == 1
    assert stats.changed
    assert "1 hidden shape(s) revealed" in stats.summary()


def test_slide_one_renders_final_state(flattened):
    out, _ = flattened
    xml = read(out, "ppt/slides/slide1.xml")
    assert etree.fromstring(xml).find(P + "timing") is None
    shapes = shapes_by_name(xml)
    assert "EXIT GONE" not in shapes, "shape whose last main-sequence effect is an exit must be removed"
    for name in ("ALWAYS VISIBLE", "HIDDEN SHAPE", "ENTRANCE STAYS", "CLICK TRIGGERED", "PARAGRAPH EXIT"):
        assert name in shapes, name
    assert shapes["HIDDEN SHAPE"].get("hidden") is None
    # the first paragraph of PARAGRAPH EXIT has exited; the second is still there
    assert "PARAGRAPH EXIT" not in texts(xml)
    assert "second paragraph" in texts(xml)


def test_slide_two_loses_transition_and_emphasis_only(flattened):
    out, _ = flattened
    xml = read(out, "ppt/slides/slide2.xml")
    root = etree.fromstring(xml)
    assert root.find(P + "timing") is None
    assert root.find(P + "transition") is None
    assert "EMPHASIS STAYS" in shapes_by_name(xml)


def test_untouched_parts_are_copied_verbatim(flattened):
    out, _ = flattened
    with zipfile.ZipFile(FIXTURE) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if not name.startswith("ppt/slides/slide"):
                assert a.read(name) == b.read(name), name


def test_namespace_prefixes_survive(flattened):
    out, _ = flattened
    xml = read(out, "ppt/slides/slide1.xml")
    assert xml.startswith(b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>")
    assert b"<p:sld " in xml and b"<a:t>" in xml
    assert b"ns0:" not in xml


def test_flattening_twice_is_a_no_op(flattened, tmp_path):
    out, _ = flattened
    stats = flatten_pptx(out, tmp_path / "again.pptx")
    assert not stats.changed
    assert stats.summary() == "nothing to flatten"


def test_bad_package_raises(tmp_path):
    bogus = tmp_path / "bogus.pptx"
    bogus.write_bytes(b"this is not a zip")
    with pytest.raises(FlattenError):
        flatten_pptx(bogus, tmp_path / "out.pptx")


# --------------------------------------------------------------------------- #
# split mode: one slide per click                                              #
# --------------------------------------------------------------------------- #


@pytest.fixture(scope="module")
def split(tmp_path_factory):
    out = tmp_path_factory.mktemp("split") / "animated.pptx"
    stats = flatten_pptx(FIXTURE, out, split_steps=True)
    return out, stats


def test_split_adds_one_slide_per_click(split):
    out, stats = split
    assert stats.slides_added == 3  # slide 1 has three clicks -> 4 states; slide 2 has none
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
    assert {n for n in names if n.startswith("ppt/slides/")} >= {
        "ppt/slides/slide3.xml", "ppt/slides/slide4.xml", "ppt/slides/slide5.xml",
        "ppt/slides/_rels/slide3.xml.rels",
    }


def test_split_pages_show_progressive_build(split):
    out, _ = split
    page = {n: read(out, f"ppt/slides/slide{n}.xml") for n in (1, 3, 4, 5, 2)}
    always = {"ALWAYS VISIBLE", "HIDDEN SHAPE", "CLICK TRIGGERED"}

    assert set(shapes_by_name(page[1])) == always
    assert set(shapes_by_name(page[3])) == always | {"ENTRANCE STAYS", "EXIT GONE"}
    assert set(shapes_by_name(page[4])) == always | {"ENTRANCE STAYS", "EXIT GONE", "PARAGRAPH EXIT"}
    assert "PARAGRAPH EXIT" in texts(page[4]) and "second paragraph" in texts(page[4])
    assert set(shapes_by_name(page[5])) == always | {"ENTRANCE STAYS", "PARAGRAPH EXIT"}
    assert "PARAGRAPH EXIT" not in texts(page[5]) and "second paragraph" in texts(page[5])
    assert set(shapes_by_name(page[2])) == {"EMPHASIS STAYS"}

    for xml in page.values():
        assert etree.fromstring(xml).find(P + "timing") is None


def test_split_registers_slides_in_order(split):
    out, _ = split
    pres = etree.fromstring(read(out, "ppt/presentation.xml"))
    rels = etree.fromstring(read(out, "ppt/_rels/presentation.xml.rels"))
    types = etree.fromstring(read(out, "[Content_Types].xml"))

    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    order = [targets[s.get(R + "id")] for s in pres.find(P + "sldIdLst")]
    assert order == ["slides/slide1.xml", "slides/slide3.xml", "slides/slide4.xml", "slides/slide5.xml", "slides/slide2.xml"]

    ids = [s.get("id") for s in pres.find(P + "sldIdLst")]
    assert len(set(ids)) == len(ids) and all(int(i) >= 256 for i in ids)
    overrides = {o.get("PartName") for o in types if o.get("PartName")}
    assert {"/ppt/slides/slide3.xml", "/ppt/slides/slide4.xml", "/ppt/slides/slide5.xml"} <= overrides


def test_split_without_animations_changes_nothing(flattened, tmp_path):
    out, _ = flattened
    stats = flatten_pptx(out, tmp_path / "again.pptx", split_steps=True)
    assert not stats.changed


# --------------------------------------------------------------------------- #
# animation replay on hand-written timing trees                               #
# --------------------------------------------------------------------------- #

TIMING = """
<p:timing xmlns:p="{ns}"><p:tnLst><p:par><p:cTn nodeType="tmRoot"><p:childTnLst>
  <p:seq><p:cTn nodeType="mainSeq"><p:childTnLst>
    {main}
  </p:childTnLst></p:cTn></p:seq>
  <p:seq><p:cTn nodeType="interactiveSeq"><p:childTnLst>
    {interactive}
  </p:childTnLst></p:cTn></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>
"""


def effect(cls: str, spid: str, paragraph: int | None = None, node: str = "clickEffect") -> str:
    inner = f'<p:txEl><p:pRg st="{paragraph}" end="{paragraph}"/></p:txEl>' if paragraph is not None else ""
    return (f'<p:par><p:cTn presetClass="{cls}" nodeType="{node}"><p:childTnLst><p:set><p:cBhvr>'
            f'<p:tgtEl><p:spTgt spid="{spid}">{inner}</p:spTgt></p:tgtEl></p:cBhvr></p:set>'
            f"</p:childTnLst></p:cTn></p:par>")


def click(*effects: str) -> str:
    return "<p:par><p:cTn><p:childTnLst>" + "".join(effects) + "</p:childTnLst></p:cTn></p:par>"


def steps(main: str, interactive: str = "") -> list[State]:
    xml = TIMING.format(ns=flatten.P_NS, main=main, interactive=interactive)
    return build_steps(etree.fromstring(xml))


def hidden(*spids: str, paragraphs=()) -> State:
    return State(frozenset(spids), frozenset(paragraphs))


def test_entrance_then_exit():
    assert steps(click(effect("entr", "1")) + click(effect("exit", "1"))) == [hidden("1"), hidden(), hidden("1")]


def test_exit_only_shape_starts_visible():
    assert steps(click(effect("exit", "1"))) == [hidden(), hidden("1")]


def test_final_state_is_last_step():
    assert steps(click(effect("exit", "1")) + click(effect("entr", "1")))[-1] == hidden()


def test_with_previous_effects_fold_into_the_same_click():
    main = click(effect("entr", "1"), effect("entr", "2", node="withEffect")) + click(effect("entr", "3"))
    assert steps(main) == [hidden("1", "2", "3"), hidden("3"), hidden()]


def test_automatic_first_group_has_no_page_of_its_own():
    main = click(effect("entr", "1", node="afterEffect")) + click(effect("entr", "2"))
    assert steps(main) == [hidden("2"), hidden()]


def test_emphasis_and_interactive_sequences_add_nothing():
    assert steps(click(effect("emph", "1"))) == [hidden()]
    assert steps("", interactive=click(effect("exit", "1"))) == [hidden()]
    assert steps(click(effect("entr", "1")), interactive=click(effect("exit", "1"))) == [hidden("1"), hidden()]


def test_paragraph_builds_track_each_paragraph():
    main = click(effect("entr", "7", paragraph=0)) + click(effect("entr", "7", paragraph=1))
    assert steps(main) == [
        hidden(paragraphs=[("7", 0), ("7", 1)]),
        hidden(paragraphs=[("7", 1)]),
        hidden(),
    ]


def test_layouts_keep_hidden_shapes_but_lose_timing():
    xml = (f'<p:sldLayout xmlns:p="{flatten.P_NS}"><p:cSld><p:spTree>'
           '<p:sp><p:nvSpPr><p:cNvPr id="2" name="guide" hidden="1"/></p:nvSpPr></p:sp>'
           '</p:spTree></p:cSld><p:timing/></p:sldLayout>')
    stats = FlattenStats()
    out = flatten_part(xml.encode(), is_slide=False, stats=stats)
    assert stats.timings_removed == 1 and stats.shapes_unhidden == 0
    assert b'hidden="1"' in out and b"<p:timing" not in out
