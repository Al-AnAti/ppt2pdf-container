# ppt2pdf-container

[![Pipeline](https://github.com/Al-AnAti/ppt2pdf-container/actions/workflows/ci.yml/badge.svg)](https://github.com/Al-AnAti/ppt2pdf-container/actions/workflows/ci.yml)

A Dockerized PowerPoint-to-PDF converter using Python and LibreOffice.

The goal is simple: convert `.ppt` and `.pptx` files to PDF without requiring Python or LibreOffice to be installed on the host machine.

## How It Works

At a high level, the container does the following:

1. Receives a file or directory path.
2. Checks for supported PowerPoint files.
3. Starts LibreOffice in headless mode.
4. Converts the presentations to PDF.
5. Places the generated PDFs next to the original files.

The host machine does not need to know anything about the Python or LibreOffice setup inside the container.

## Why Docker?

The actual conversion is handled by LibreOffice through `soffice --headless`. Since LibreOffice is a system-level dependency, installing this as a normal Python package would still require users to install and configure LibreOffice separately.

Docker makes this easier by packaging everything together:

* Python
* LibreOffice
* The conversion script

This means the host machine only needs Docker, and the same environment can be used wherever the container runs.

## Features

* **No Python or LibreOffice installation required** — only Docker is needed.
* **Single file conversion** — convert a specific `.ppt` or `.pptx` file.
* **Batch conversion** — pass a directory to convert all supported presentations inside it.
* **Automatic output** — PDFs are created in the same directory as the original files.
* **Reproducible environment** — the conversion dependencies are packaged inside the Docker image.

## Usage

To avoid typing out long absolute paths every time, you can drop a `docker-compose.yml` file into the folder where your presentations live:

~~~yaml
services:
  converter:
    image: nooralanati/ppt2pdf-container:latest
    volumes:
      - .:/data
    working_dir: /data
~~~

Once that file is in your folder, open your terminal in that folder and run:

**For a single file:**
~~~bash
docker compose run --rm converter presentation.pptx
~~~

**For batch conversion (all files in the folder):**
~~~bash
docker compose run --rm converter .
~~~
