"""Unit tests for the ppt2pdf CLI helpers (no LibreOffice needed)."""
import os
import time

import pytest

import ppt2pdf


def touch(path, mtime=None):
    path.write_bytes(b"x")
    if mtime is not None:
        os.utime(path, (mtime, mtime))


def test_defaults_to_current_directory():
    args = ppt2pdf.parse_args([])
    assert args.path == "."
    assert not args.recursive and not args.force and not args.no_flatten and not args.split_steps
    assert args.jobs is None
    assert args.timeout == ppt2pdf.DEFAULT_TIMEOUT


def test_short_flags():
    args = ppt2pdf.parse_args(["-rfs", "-j", "3", "decks"])
    assert args.recursive and args.force and args.split_steps and args.jobs == 3 and args.path == "decks"


def test_split_steps_conflicts_with_no_flatten():
    with pytest.raises(SystemExit):
        ppt2pdf.parse_args(["--split-steps", "--no-flatten"])


@pytest.mark.parametrize("argv", [["-j", "0"], ["--timeout", "0"]])
def test_rejects_bad_numbers(argv):
    with pytest.raises(SystemExit):
        ppt2pdf.parse_args(argv)


def test_discover_single_file_and_rejects_other_types(tmp_path):
    deck = tmp_path / "deck.PPTX"
    touch(deck)
    assert ppt2pdf.discover(deck, recursive=False) == [deck]
    with pytest.raises(SystemExit):
        ppt2pdf.discover(tmp_path / "notes.txt", recursive=False)
    with pytest.raises(SystemExit):
        ppt2pdf.discover(tmp_path / "missing.pptx", recursive=False)


def test_discover_directory_and_recursion(tmp_path):
    touch(tmp_path / "b.pptx")
    touch(tmp_path / "a.ppt")
    touch(tmp_path / "~$a.ppt")  # Office lock file
    touch(tmp_path / "readme.pdf")
    (tmp_path / "nested" / "deeper").mkdir(parents=True)
    touch(tmp_path / "nested" / "deeper" / "c.ppsx")

    flat = ppt2pdf.discover(tmp_path, recursive=False)
    assert [p.name for p in flat] == ["a.ppt", "b.pptx"]

    deep = ppt2pdf.discover(tmp_path, recursive=True)
    assert [p.name for p in deep] == ["a.ppt", "b.pptx", "c.ppsx"]


def test_up_to_date_compares_mtimes(tmp_path):
    src = tmp_path / "deck.pptx"
    pdf = tmp_path / "deck.pdf"
    now = time.time()
    touch(src, now)
    assert not ppt2pdf.is_up_to_date(src, pdf)
    touch(pdf, now - 10)
    assert not ppt2pdf.is_up_to_date(src, pdf)
    touch(pdf, now + 10)
    assert ppt2pdf.is_up_to_date(src, pdf)


def test_main_reports_empty_directory(tmp_path, capsys):
    assert ppt2pdf.main([str(tmp_path)]) == 0
    assert "no presentations found" in capsys.readouterr().out


def test_main_skips_everything_without_launching_workers(tmp_path, capsys):
    src = tmp_path / "deck.pptx"
    touch(src)
    touch(tmp_path / "deck.pdf", time.time() + 10)
    assert ppt2pdf.main([str(tmp_path)]) == 0
    out = capsys.readouterr().out
    assert "[1/1] SKIP deck.pptx" in out
    assert "0 converted, 1 skipped" in out
