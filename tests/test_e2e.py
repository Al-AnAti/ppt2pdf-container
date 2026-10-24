"""End-to-end tests that run the built image. Skipped unless PPT2PDF_IMAGE is set:

    docker build -t ppt2pdf-test .
    PPT2PDF_IMAGE=ppt2pdf-test pytest tests/test_e2e.py
"""
import os
import shutil
import subprocess
import time
from pathlib import Path

import pytest
from pypdf import PdfReader

IMAGE = os.environ.get("PPT2PDF_IMAGE")
FIXTURE = Path(__file__).parent / "fixtures" / "animated.pptx"

pytestmark = pytest.mark.skipif(not IMAGE, reason="PPT2PDF_IMAGE not set")


def run(workdir: Path, *args: str, entrypoint: str | None = None) -> subprocess.CompletedProcess:
    cmd = ["docker", "run", "--rm", "--init", "-v", f"{workdir}:/data"]
    if entrypoint:
        cmd += ["--entrypoint", entrypoint]
    cmd += [IMAGE, *args]
    return subprocess.run(cmd, capture_output=True, text=True, timeout=600)


def pdf_text(path: Path) -> str:
    return "\n".join(page.extract_text() for page in PdfReader(str(path)).pages)


@pytest.fixture
def workdir(tmp_path):
    shutil.copy(FIXTURE, tmp_path / "animated.pptx")
    return tmp_path


def test_flattened_pdf_shows_final_state(workdir):
    proc = run(workdir, "-v")
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "[1/1] OK" in proc.stdout and "1 converted" in proc.stdout
    assert "flattened:" in proc.stdout
    assert "convert /data" not in proc.stdout, "raw LibreOffice output must not leak"

    text = pdf_text(workdir / "animated.pdf")
    for expected in ("ALWAYS VISIBLE", "HIDDEN SHAPE", "ENTRANCE STAYS", "CLICK TRIGGERED", "second paragraph", "EMPHASIS STAYS"):
        assert expected in text, expected
    assert "EXIT GONE" not in text
    assert "PARAGRAPH EXIT" not in text, "a paragraph that has exited is blanked"
    assert len(PdfReader(str(workdir / "animated.pdf")).pages) == 2, "one page per slide by default"


def test_split_steps_writes_one_page_per_click(workdir):
    proc = run(workdir, "--split-steps", "-v")
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "3 step slide(s) added" in proc.stdout
    pages = [page.extract_text() for page in PdfReader(str(workdir / "animated.pdf")).pages]
    assert len(pages) == 5

    assert "ENTRANCE STAYS" not in pages[0] and "HIDDEN SHAPE" in pages[0] and "CLICK TRIGGERED" in pages[0]
    assert "ENTRANCE STAYS" in pages[1] and "EXIT GONE" in pages[1] and "PARAGRAPH EXIT" not in pages[1]
    assert "PARAGRAPH EXIT" in pages[2] and "second paragraph" in pages[2]
    assert "EXIT GONE" not in pages[3] and "PARAGRAPH EXIT" not in pages[3] and "second paragraph" in pages[3]
    assert "EMPHASIS STAYS" in pages[4]


def test_no_flatten_keeps_libreoffice_behaviour(workdir):
    proc = run(workdir, "--no-flatten")
    assert proc.returncode == 0, proc.stdout + proc.stderr
    text = pdf_text(workdir / "animated.pdf")
    assert "HIDDEN SHAPE" not in text
    assert "EXIT GONE" in text


def test_skip_and_force(workdir):
    assert run(workdir).returncode == 0
    first = (workdir / "animated.pdf").stat().st_mtime

    again = run(workdir)
    assert again.returncode == 0
    assert "SKIP" in again.stdout and "1 skipped" in again.stdout
    assert (workdir / "animated.pdf").stat().st_mtime == first

    time.sleep(1.1)
    forced = run(workdir, "--force")
    assert forced.returncode == 0
    assert "OK" in forced.stdout and "1 converted" in forced.stdout
    assert (workdir / "animated.pdf").stat().st_mtime > first


def test_recursive_and_legacy_ppt(workdir):
    nested = workdir / "nested" / "deeper"
    nested.mkdir(parents=True)
    # Let LibreOffice itself produce a binary .ppt copy of the fixture.
    conv = run(workdir, "--headless", "--convert-to", "ppt", "--outdir", "/data/nested/deeper", "/data/animated.pptx", entrypoint="soffice")
    assert (nested / "animated.ppt").exists(), conv.stdout + conv.stderr

    shallow = run(workdir)
    assert "1 presentation in" in shallow.stdout
    assert not (nested / "animated.pdf").exists()

    deep = run(workdir, "-r", "-f", "-v")
    assert deep.returncode == 0, deep.stdout + deep.stderr
    assert "2 presentations in" in deep.stdout and "2 converted" in deep.stdout
    assert "[2/2]" in deep.stdout
    text = pdf_text(nested / "animated.pdf")
    assert "HIDDEN SHAPE" in text and "EXIT GONE" not in text


def test_failure_is_reported_cleanly(workdir):
    (workdir / "broken.pptx").write_bytes(b"not really a presentation")
    proc = run(workdir)
    assert proc.returncode == 1
    assert "FAIL broken.pptx" in proc.stdout
    assert "1 converted, 1 failed" in proc.stdout
    assert not (workdir / "broken.pdf").exists()
