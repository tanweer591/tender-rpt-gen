# ─────────────────────────────────────────────────────────────────────────────
#  Aerotender — Defence Tender Scraper
#  Dockerfile for Railway deployment
#  Base: Python 3.11 slim + Chromium (headless) + ChromeDriver
# ─────────────────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# ── System dependencies ───────────────────────────────────────────────────────
RUN apt-get update && apt-get install -y --no-install-recommends \
        chromium \
        chromium-driver \
        fonts-liberation \
        libnss3 \
        libatk-bridge2.0-0 \
        libgtk-3-0 \
        libxss1 \
        libgbm1 \
        libasound2 \
        wget \
        ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# ── Tell aerotender to use the system Chrome/Chromedriver ────────────────────
ENV CHROME_BIN=/usr/bin/chromium
ENV CHROMEDRIVER_PATH=/usr/bin/chromedriver

# Chromedriver needs a writable user-data dir
ENV CHROME_USER_DATA=/tmp/chrome-data

# ── App directory ─────────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies ───────────────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Application source ────────────────────────────────────────────────────────
COPY . .

# Railway sets $PORT; gunicorn binds to it.
# Two workers, each handling one request at a time.
# Timeout raised to 600 s because scraping can take several minutes.
CMD gunicorn app:app \
        --bind 0.0.0.0:${PORT:-8000} \
        --workers 1 \
        --threads 4 \
        --timeout 600 \
        --log-level info
