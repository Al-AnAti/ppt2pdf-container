# ppt2pdf-container

[![Pipeline](https://github.com/Al-AnAti/ppt2pdf-container/actions/workflows/ci.yml/badge.svg)](https://github.com/Al-AnAti/ppt2pdf-container/actions/workflows/ci.yml)

A Dockerized PowerPoint-to-PDF converter using Python and LibreOffice.

The goal is simple: convert `.ppt` and `.pptx` files to PDF without requiring Python or LibreOffice to be installed on the host machine, and get PDFs that look like the slides do once they have finished playing.

## How It Works

For every presentation the container:

1. **Flattens it to its final visible state.** Every animation block is stripped, shapes that are hidden in the file (the "eye" toggle in PowerPoint's selection pane) are revealed, click-triggered objects are kept, and the main animation sequence is replayed so shapes or paragraphs that have exited are gone. Slide transitions are dropped too. The result is exactly one page per slide. With `--split-steps` the slide is instead written once per click, like a printed build. Legacy `.ppt` files are re-saved as `.pptx` by LibreOffice first so they get the same treatment.
2. **Converts it with its own headless LibreOffice instance** in a private scratch directory, so several files can convert at the same time.
3. **Moves the finished PDF next to the source file**, replacing any old one atomically.

Files whose PDF is already newer than the presentation are skipped unless you pass `--force`.

## Why Docker?

The actual conversion is handled by LibreOffice through `soffice --headless`. Since LibreOffice is a system-level dependency, installing this as a normal Python package would still require users to install and configure LibreOffice separately.

Docker packages everything together: Python, LibreOffice, the conversion script and a font set chosen so that text keeps its layout. The host machine only needs Docker.

## Features

* **Static final state** — entrance, exit and emphasis animations are stripped; hidden and click-triggered objects render; objects that have exited are gone. One PDF page per slide, always. Pass `--no-flatten` to convert files exactly as they are.
* **Optional step splitting** — `--split-steps` / `-s` writes one page per animation click: the slide before the first click, then after each click, with bullet-by-bullet builds included.
* **Smart skipping** — a presentation is skipped when its PDF is already newer. `--force` / `-f` re-converts everything.
* **File or directory input** — pass a single file, a directory, or nothing at all (defaults to the current directory). `--recursive` / `-r` walks subdirectories, writing each PDF alongside its source.
* **Parallel conversion** — files are spread over a process pool, one LibreOffice per worker. `--jobs N` / `-j N` sets the worker count (default: number of CPUs).
* **Font parity** — the image ships Liberation (Arial / Times New Roman / Courier New), Carlito (Calibri), Caladea (Cambria), URW base35 (Palatino, Bookman, Century Schoolbook), DejaVu and Noto (including CJK and emoji), with fontconfig aliases for Segoe UI, Aptos, Tahoma, Verdana and other common Office faces that have no metric-compatible clone.
* **Clean progress output** — one line per file with `[n/total]`, status, elapsed time, then a summary. LibreOffice's own chatter is captured and only shown (last lines) when a conversion fails. `-v` adds what was flattened in each file.
* **Supported inputs** — `.pptx`, `.pptm`, `.ppsx`, `.potx`, `.ppt`, `.pps`, `.pot`. Office lock files (`~$deck.pptx`) are ignored.

## Usage

The main way to use this project is the `ppt2pdf` PowerShell function in [`scripts/ppt2pdf.ps1`](scripts/ppt2pdf.ps1). Once it is in your profile you can run `ppt2pdf` from any folder, on any file or directory, with no Docker commands to remember. The manual Docker Compose route further down works everywhere else.

### 1. Install the `ppt2pdf` command (Windows, PowerShell)

Clone the repository, then run these commands once in PowerShell. Edit the path in step 2 to point at your clone.

~~~powershell
# 1. create the profile file if it does not exist yet
if (-not (Test-Path $PROFILE)) { New-Item -ItemType File -Path $PROFILE -Force | Out-Null }

# 2. load the script from your clone every time PowerShell starts
Add-Content $PROFILE '. "C:\path\to\ppt2pdf-container\scripts\ppt2pdf.ps1"'

# 3. reload the profile in the current session (new terminals pick it up automatically)
. $PROFILE
~~~

If you would rather not keep the clone around, copy the function itself into the profile instead of step 2:

~~~powershell
Get-Content "C:\path\to\ppt2pdf-container\scripts\ppt2pdf.ps1" | Add-Content $PROFILE
~~~

If PowerShell refuses to load the profile ("running scripts is disabled on this system"), allow local scripts for your user once and open a new terminal:

~~~powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
~~~

To look at or edit the profile by hand: `notepad $PROFILE`.

### 2. Use it

~~~powershell
ppt2pdf                       # everything in the current folder
ppt2pdf .                     # same thing
ppt2pdf "D:\Courses\Deck"     # a folder
ppt2pdf .\lecture.pptx        # one file (its folder is mounted, the file name is passed through)
ppt2pdf -r -f                 # recurse and re-convert everything
ppt2pdf -SplitSteps           # one page per click instead of the final state only
ppt2pdf -Jobs 2 -Verbose      # two workers, show what was flattened
ppt2pdf -WhatIf               # print the docker command instead of running it
~~~

The function resolves any Windows path, mounts the right folder into the container and forwards the flags. It checks that the `docker` command exists and that the daemon is running before it starts, runs the container with `--rm --init`, and attaches a TTY only when you are in a real terminal so colours work interactively and output stays plain when redirected. `$LASTEXITCODE` is the container's exit status afterwards (`0` all good, `1` some file failed, `2` the wrapper refused to start).

Set `$env:PPT2PDF_IMAGE` (or pass `-Image`) to run a locally built image instead of the one from Docker Hub. Docker only pulls an image it does not have locally, so after `docker build -t nooralanati/ppt2pdf-container:latest .` the function and the compose file use your local build without any extra configuration.

### Running it manually with Docker Compose

On macOS, Linux, or without the PowerShell function, drop a `docker-compose.yml` file into the folder where your presentations live:

~~~yaml
services:
  converter:
    image: nooralanati/ppt2pdf-container:latest
    init: true
    volumes:
      - .:/data
    working_dir: /data
~~~

Then open a terminal in that folder and run:

**Everything in the folder (the default):**
~~~bash
docker compose run --rm converter
~~~

**A single file:**
~~~bash
docker compose run --rm converter presentation.pptx
~~~

**Every presentation in the folder and its subfolders, re-converting even up-to-date ones:**
~~~bash
docker compose run --rm converter -rf
~~~

**One page per animation click:**
~~~bash
docker compose run --rm converter --split-steps
~~~

**Without Compose at all:**
~~~bash
docker run --rm --init -v "$PWD:/data" nooralanati/ppt2pdf-container:latest -r
~~~

### Options

~~~text
usage: ppt2pdf [-h] [-r] [-f] [-j N] [--timeout SEC] [-s] [--no-flatten] [-v] [--version] [path]

positional arguments:
  path            a presentation or a directory of them (default: current directory)

options:
  -r, --recursive also convert presentations in subdirectories
  -f, --force     re-convert even if an up-to-date PDF already exists
  -j, --jobs N    parallel LibreOffice workers (default: CPU count)
  --timeout SEC   per-file LibreOffice timeout (default: 300)
  -s, --split-steps
                  write one PDF page per animation click instead of only the final state of each slide
  --no-flatten    convert as-is: keep animations, leave hidden shapes hidden
  -v, --verbose   show what was flattened in each file
~~~

Exit status is `0` when every file converted or was skipped, `1` when any file failed.

### Example output

~~~text
ppt2pdf: 4 presentations in . (recursive), 3 to convert with 3 workers
[1/4] SKIP archive/q1.pptx  (PDF is up to date)
[2/4] OK   intro.pptx -> intro.pdf  (6.8s)
[3/4] OK   archive/q2.ppt -> q2.pdf  (11.2s)
[4/4] FAIL broken.pptx  (0.9s)
      Error: source file could not be loaded
Done in 12.0s: 2 converted, 1 skipped, 1 failed
~~~

## What "final visible state" means

PowerPoint decks often rely on animations: text that appears on click, objects that fly out, builds that reveal bullets one by one. A static render has to pick one moment. By default this tool picks the end of the slide and writes one page per slide; with `--split-steps` it writes every moment a click would reveal.

| In the deck | Default (final state) | `--split-steps` |
|---|---|---|
| Entrance animation | Shown | Hidden until its click |
| Exit animation | Removed (shape) or blanked (paragraph) | Shown until its click |
| Emphasis or motion-path animation | Shown, no effect | Shown, no extra page |
| Effects set to "with/after previous" | Part of the same state | Part of the same click |
| Hidden shape (`hidden="1"`) | Shown | Shown |
| Shape revealed by clicking another shape (interactive sequence) | Shown | Shown |
| Chart-series or character-range effects | Not modelled, shape shown | Not modelled, shape shown |
| Hidden slide | Not exported by LibreOffice | Not exported by LibreOffice |

## Development

~~~bash
pip install -r requirements-dev.txt
pytest tests/test_flatten.py tests/test_cli.py   # fast, no LibreOffice needed
docker build -t ppt2pdf-test .
PPT2PDF_IMAGE=ppt2pdf-test pytest tests/test_e2e.py   # runs the image against tests/fixtures
~~~

`tests/fixtures/animated.pptx` is generated by `tests/fixtures/make_fixture.py` and contains one shape for every case the flattener handles.
