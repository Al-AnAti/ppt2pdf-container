FROM python:3.11-slim

RUN apt-get update && \
apt-get install -y --no-install-recommends libreoffice && \
apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /usr/src/app

COPY ppt2pdf.py .

ENTRYPOINT ["python", "ppt2pdf.py"]