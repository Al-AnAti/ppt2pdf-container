#!/usr/bin/env python3
"""ppt2pdf - convert PowerPoint presentations to PDF with headless LibreOffice.

Each deck is flattened first (animations stripped, hidden shapes revealed,
and either the final state kept or, with --split-steps, one slide written per
click), converted into a private
scratch directory by its own LibreOffice instance, and the finished PDF is
moved next to the source file. Files are processed in parallel by a process
pool; PDFs that are already newer than their source are skipped.
"""
from __future__ import annotations

import argparse
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from concurrent.futures import FIRST_COMPLETED, Future, ProcessPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path

from flatten import FlattenError, flatten_pptx

__version__ = "2.0.0"

OOXML_EXTS = {".pptx", ".pptm", ".ppsx", ".potx"}
BINARY_EXTS = {".ppt", ".pps", ".pot"}
SUPPORTED_EXTS = OOXML_EXTS | BINARY_EXTS
TEMP_PREFIXES = ("~$", ".~lock")  # Office / LibreOffice lock files
SOFFICE = os.environ.get("SOFFICE", "soffice")
NOISE = ("javaldx",)  # harmless LibreOffice chatter never worth showing
DEFAULT_TIMEOUT = 300.0


class ConversionError(Exception):
    pass


@dataclass(frozen=True)
class Job:
    src: Path
    pdf: Path
    flatten: bool
    split_steps: bool
    timeout: float


@dataclass
class Result:
    job: Job
    ok: bool
    seconds: float
    notes: list[str]  # flatten summary, warnings, or the error text


# --------------------------------------------------------------------------- #
# worker process                                                               #
# --------------------------------------------------------------------------- #

_WORK_ROOT: Path | None = None
_PROFILE: Path | None = None
_ENV: dict[str, str] | None = None


def _init_worker(work_root: str) -> None:
    """Give every worker its own LibreOffice profile; concurrent soffice
    instances cannot share one."""
    global _WORK_ROOT, _PROFILE, _ENV
    _WORK_ROOT = Path(work_root)
    _PROFILE = _WORK_ROOT / f"profile-{os.getpid()}"
    _PROFILE.mkdir(parents=True, exist_ok=True)
    # Also give it a private HOME so instances never share ~/.config or ~/.cache.
    # (LibreOffice can still crash at startup when several instances launch at
    # once in a fresh container; run_soffice retries that once.)
    home = _PROFILE / "home"
    home.mkdir(exist_ok=True)
    _ENV = {**os.environ, "HOME": str(home), "XDG_CONFIG_HOME": str(home / ".config"), "XDG_CACHE_HOME": str(home / ".cache")}


def _kill(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    try:
        if hasattr(os, "killpg"):
            os.killpg(proc.pid, signal.SIGKILL)  # soffice forks helpers; kill the whole session
        else:
            proc.kill()
    except ProcessLookupError:
        pass


def _tail(text: str, lines: int = 6) -> str:
    kept = [ln.strip() for ln in text.splitlines() if ln.strip() and not any(n in ln for n in NOISE)]
    return "\n".join(kept[-lines:])


def run_soffice(src: Path, fmt: str, outdir: Path, timeout: float, notes: list[str] | None = None) -> Path:
    """Convert ``src`` to ``fmt`` into ``outdir`` and return the produced file.

    A crashed LibreOffice (as opposed to a timeout or an unreadable file) is
    retried once; the retry is recorded in ``notes``.
    """
    try:
        return _run_soffice_once(src, fmt, outdir, timeout)
    except ConversionError as exc:
        if "timed out" in str(exc) or "could not be loaded" in str(exc) or "not found" in str(exc):
            raise
        if notes is not None:
            notes.append(f"warning: LibreOffice crashed once ({_tail(str(exc), 1)}); retried")
        return _run_soffice_once(src, fmt, outdir, timeout)


def _run_soffice_once(src: Path, fmt: str, outdir: Path, timeout: float) -> Path:
    assert _PROFILE is not None, "run_soffice called outside a worker"
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        SOFFICE,
        f"-env:UserInstallation={_PROFILE.resolve().as_uri()}",
        "--headless",
        "--norestore",
        "--convert-to", fmt,
        "--outdir", str(outdir),
        str(src),
    ]
    popen_kwargs = {"start_new_session": True} if hasattr(os, "killpg") else {}
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, errors="replace", stdin=subprocess.DEVNULL, env=_ENV, **popen_kwargs,
        )
    except FileNotFoundError:
        raise ConversionError(f"LibreOffice not found (looked for '{SOFFICE}')")
    try:
        output, _ = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        _kill(proc)
        proc.communicate()
        raise ConversionError(f"LibreOffice timed out after {timeout:.0f}s")
    except BaseException:
        _kill(proc)
        raise

    produced = sorted(outdir.glob(f"*.{fmt}"))
    if proc.returncode != 0 or not produced:
        reason = _tail(output) or f"LibreOffice exited with status {proc.returncode} and produced no {fmt}"
        raise ConversionError(reason)
    return produced[0]


def _publish(tmp_pdf: Path, final_pdf: Path) -> None:
    """Copy next to the source, then rename atomically so a half-written PDF
    never masquerades as an up-to-date one."""
    partial = final_pdf.with_name(final_pdf.name + ".part")
    try:
        shutil.copyfile(tmp_pdf, partial)
        os.replace(partial, final_pdf)
    finally:
        partial.unlink(missing_ok=True)


def convert(job: Job) -> Result:
    started = time.perf_counter()
    notes: list[str] = []
    try:
        with tempfile.TemporaryDirectory(dir=_WORK_ROOT, prefix="job-") as tmp_name:
            tmp = Path(tmp_name)
            src = job.src
            if job.flatten:
                if src.suffix.lower() in BINARY_EXTS:
                    # The binary .ppt format has no XML to edit; let LibreOffice
                    # re-save it as .pptx first, then flatten that.
                    src = run_soffice(src, "pptx", tmp / "ooxml", job.timeout, notes)
                flat = tmp / "flat" / (job.src.stem + ".pptx")
                flat.parent.mkdir()
                try:
                    stats = flatten_pptx(src, flat, split_steps=job.split_steps)
                    notes.append(f"flattened: {stats.summary()}")
                    src = flat
                except FlattenError as exc:
                    notes.append(f"warning: could not flatten ({exc}); converted as-is")
                    src = job.src
            pdf = run_soffice(src, "pdf", tmp / "out", job.timeout, notes)
            _publish(pdf, job.pdf)
        return Result(job, True, time.perf_counter() - started, notes)
    except ConversionError as exc:
        notes.append(str(exc))
    except Exception as exc:  # noqa: BLE001 - report, never crash the pool
        notes.append(f"{type(exc).__name__}: {exc}")
    return Result(job, False, time.perf_counter() - started, notes)


# --------------------------------------------------------------------------- #
# main process                                                                 #
# --------------------------------------------------------------------------- #


class Console:
    """Minimal coloured status output; colour only when writing to a terminal."""

    STYLES = {"OK": "32", "SKIP": "36", "FAIL": "31", "WARN": "33"}

    def __init__(self, stream=None):
        self.stream = stream or sys.stdout
        self.color = self.stream.isatty() and not os.environ.get("NO_COLOR")

    def tag(self, kind: str) -> str:
        text = f"{kind:<4}"
        return f"\033[1;{self.STYLES[kind]}m{text}\033[0m" if self.color else text

    def line(self, text: str = "") -> None:
        print(text, file=self.stream, flush=True)

    def status(self, index: int, total: int, kind: str, text: str, notes: list[str] = ()) -> None:
        width = len(str(total))
        self.line(f"[{index:>{width}}/{total}] {self.tag(kind)} {text}")
        for note in notes:
            for ln in note.splitlines():
                self.line(" " * (width * 2 + 4) + ln)


def discover(target: Path, recursive: bool) -> list[Path]:
    if target.is_file():
        if target.suffix.lower() not in SUPPORTED_EXTS:
            raise SystemExit(f"error: '{target}' is not a PowerPoint file ({', '.join(sorted(SUPPORTED_EXTS))})")
        return [target]
    if not target.is_dir():
        raise SystemExit(f"error: '{target}' does not exist")
    candidates = target.rglob("*") if recursive else target.iterdir()
    return sorted(
        p for p in candidates
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS and not p.name.startswith(TEMP_PREFIXES)
    )


def is_up_to_date(src: Path, pdf: Path) -> bool:
    try:
        return pdf.stat().st_mtime >= src.stat().st_mtime
    except FileNotFoundError:
        return False


def fmt_seconds(seconds: float) -> str:
    return f"{seconds:.1f}s" if seconds < 60 else f"{int(seconds // 60)}m{seconds % 60:04.1f}s"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="ppt2pdf",
        description="Convert PowerPoint presentations (.ppt/.pptx) to PDF with headless LibreOffice. "
                    "PDFs are written next to their source files.",
    )
    parser.add_argument("path", nargs="?", default=".", help="a presentation or a directory of them (default: current directory)")
    parser.add_argument("-r", "--recursive", action="store_true", help="also convert presentations in subdirectories")
    parser.add_argument("-f", "--force", action="store_true", help="re-convert even if an up-to-date PDF already exists")
    parser.add_argument("-j", "--jobs", type=int, default=None, metavar="N", help="parallel LibreOffice workers (default: CPU count)")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, metavar="SEC", help=f"per-file LibreOffice timeout (default: {DEFAULT_TIMEOUT:.0f})")
    parser.add_argument("-s", "--split-steps", action="store_true",
                        help="write one PDF page per animation click instead of only the final state of each slide")
    parser.add_argument("--no-flatten", action="store_true", help="convert as-is: keep animations, leave hidden shapes hidden")
    parser.add_argument("-v", "--verbose", action="store_true", help="show what was flattened in each file")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)
    if args.jobs is not None and args.jobs < 1:
        parser.error("--jobs must be at least 1")
    if args.timeout <= 0:
        parser.error("--timeout must be positive")
    if args.split_steps and args.no_flatten:
        parser.error("--split-steps and --no-flatten are mutually exclusive")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    console = Console()
    target = Path(args.path)
    files = discover(target, args.recursive)
    base = target if target.is_dir() else target.parent

    def label(p: Path) -> str:
        try:
            return str(p.relative_to(base))
        except ValueError:
            return str(p)

    if not files:
        where = f"{target}{' (recursive)' if args.recursive else ''}"
        console.line(f"ppt2pdf: no presentations found in {where}")
        return 0

    total = len(files)
    jobs: list[Job] = []
    skips: list[Path] = []
    for src in files:
        pdf = src.with_suffix(".pdf")
        if not args.force and is_up_to_date(src, pdf):
            skips.append(src)
        else:
            jobs.append(Job(src=src, pdf=pdf, flatten=not args.no_flatten, split_steps=args.split_steps, timeout=args.timeout))

    workers = min(args.jobs or os.cpu_count() or 1, max(len(jobs), 1))
    noun = "presentation" if total == 1 else "presentations"
    console.line(f"ppt2pdf: {total} {noun} in {target}{' (recursive)' if args.recursive else ''}"
                 f", {len(jobs)} to convert with {workers} worker{'s' if workers != 1 else ''}")

    index = 0
    for src in skips:
        index += 1
        console.status(index, total, "SKIP", f"{label(src)}  (PDF is up to date)")

    converted = failed = 0
    started = time.perf_counter()
    if jobs:
        work_root = tempfile.mkdtemp(prefix="ppt2pdf-")
        try:
            with ProcessPoolExecutor(max_workers=workers, initializer=_init_worker, initargs=(work_root,)) as pool:
                pending: set[Future] = {pool.submit(convert, job) for job in jobs}
                try:
                    while pending:
                        done, pending = wait(pending, return_when=FIRST_COMPLETED)
                        for future in done:
                            result: Result = future.result()
                            index += 1
                            name = label(result.job.src)
                            elapsed = fmt_seconds(result.seconds)
                            if result.ok:
                                converted += 1
                                notes = result.notes if args.verbose else [n for n in result.notes if n.startswith("warning")]
                                console.status(index, total, "OK", f"{name} -> {result.job.pdf.name}  ({elapsed})", notes)
                            else:
                                failed += 1
                                console.status(index, total, "FAIL", f"{name}  ({elapsed})", result.notes)
                except KeyboardInterrupt:
                    for future in pending:
                        future.cancel()
                    console.line("\nppt2pdf: interrupted, waiting for running conversions to stop...")
                    pool.shutdown(wait=False, cancel_futures=True)
                    return 130
        finally:
            shutil.rmtree(work_root, ignore_errors=True)

    summary = [f"{converted} converted"]
    if skips:
        summary.append(f"{len(skips)} skipped")
    if failed:
        summary.append(f"{failed} failed")
    console.line(f"Done in {fmt_seconds(time.perf_counter() - started)}: {', '.join(summary)}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
