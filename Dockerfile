# Qlon web interface.
#
# Heavy by necessity: the render pipeline needs Python, Quarto, and a headless
# Chromium (via Playwright, for Mermaid diagrams). Expect a ~1.5 GB+ image.
FROM python:3.13-slim

ARG QUARTO_VERSION=1.6.40

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PLAYWRIGHT_BROWSERS_PATH=/opt/playwright

# System deps + Quarto. The Playwright OS libraries are installed later via
# `playwright install --with-deps`.
RUN apt-get update && apt-get install -y --no-install-recommends \
        wget ca-certificates \
    && dpkgArch="$(dpkg --print-architecture)" \
    && wget -qO /tmp/quarto.deb \
        "https://github.com/quarto-dev/quarto-cli/releases/download/v${QUARTO_VERSION}/quarto-${QUARTO_VERSION}-linux-${dpkgArch}.deb" \
    && dpkg -i /tmp/quarto.deb \
    && rm /tmp/quarto.deb \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps first for better layer caching. The web requirements file is
# self-contained (the repo-root requirements.txt is a Windows-only frozen dump).
COPY web/requirements-web.txt ./
RUN pip install -r requirements-web.txt \
    && playwright install --with-deps chromium

# Copy the project.
COPY . .

EXPOSE 5000

WORKDIR /app/web
# --timeout 600: a single Quarto render (with diagrams) can run well past the
# default 30s gunicorn worker timeout.
CMD ["gunicorn", "-w", "2", "--timeout", "600", "-b", "0.0.0.0:5000", "app:app"]
