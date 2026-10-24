FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

# Headless LibreOffice Impress plus fonts that stand in for what PowerPoint
# decks usually ask for, so text keeps its metrics instead of overflowing:
#   fonts-liberation          Arial, Times New Roman, Courier New (metric-compatible)
#   fonts-crosextra-carlito   Calibri (metric-compatible)
#   fonts-crosextra-caladea   Cambria (metric-compatible)
#   fonts-urw-base35          Palatino / Book Antiqua, Bookman, Century Schoolbook, Avant Garde
#   fonts-dejavu              Verdana / Tahoma stand-in, wide Latin, Greek, Cyrillic coverage
#   fonts-noto-*              Segoe UI / Aptos stand-in, CJK and emoji coverage
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress-nogui \
        fontconfig \
        fonts-liberation \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-urw-base35 \
        fonts-dejavu \
        fonts-noto-core \
        fonts-noto-cjk \
        fonts-noto-color-emoji \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Map Office fonts that have no metric-compatible clone onto the closest open font.
COPY fonts/local.conf /etc/fonts/local.conf
RUN fc-cache -f

WORKDIR /usr/src/app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY ppt2pdf.py flatten.py ./

# Mount your presentations here; the script defaults to the current directory.
WORKDIR /data
ENTRYPOINT ["python", "/usr/src/app/ppt2pdf.py"]
