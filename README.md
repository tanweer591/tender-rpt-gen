# Aerotender — Defence Tender Scraper on Railway

Scrapes **DefProc** (NIC GePNIC) and **GeM** portals in parallel using
headless Chrome, then exports results to a formatted `.xlsx` file.

---

## Files

| File | Purpose |
|---|---|
| `aerotender.py` | Core scraper (DefProc + GeM + Excel export) |
| `app.py` | Flask backend (REST API + file serving) |
| `tender_ui.html` | Browser UI (served by Flask at `/`) |
| `Dockerfile` | Installs Chrome + Python deps for Railway |
| `railway.toml` | Railway build/deploy config |
| `requirements.txt` | Python dependencies |

---

## Deploy to Railway (step-by-step)

### Option A — Deploy via GitHub (recommended)

1. **Create a GitHub repo** and push this folder to it:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```

2. **Go to [railway.app](https://railway.app)** → *New Project* → *Deploy from GitHub repo*

3. Select your repo. Railway auto-detects the `Dockerfile` and `railway.toml`.

4. Click **Deploy**. The first build takes ~3–4 minutes (Chrome download).

5. Once deployed, go to **Settings → Networking → Generate Domain** to get a public URL.

---

### Option B — Deploy via Railway CLI

```bash
# Install Railway CLI
npm install -g @railway/cli       # or: brew install railway

# Login
railway login

# Create project & deploy
railway init
railway up
```

---

## Environment Variables (set in Railway dashboard)

| Variable | Default | Notes |
|---|---|---|
| `PORT` | set by Railway | Do **not** override — Railway controls this |
| `CHROME_BIN` | `/usr/bin/chromium` | Already set in Dockerfile |
| `CHROMEDRIVER_PATH` | `/usr/bin/chromedriver` | Already set in Dockerfile |

No other secrets are needed.

---

## Local Development

```bash
# Create a virtual environment
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# Install deps (Chrome must be installed locally)
pip install -r requirements.txt

# Run the Flask dev server
python app.py
# → open http://localhost:5000
```

---

## How it works

1. Open the URL → the UI loads (`tender_ui.html`).
2. Click **Run Scraper** → POST `/run` starts a background thread.
3. The UI polls GET `/status` every 2 s and shows live progress for both portals.
4. When done, click **Download Excel** → GET `/download` streams the `.xlsx`.

---

## Notes

- Scraping takes **3–8 minutes** depending on network speed.
- Railway's free tier has enough memory (512 MB); Chrome uses ~200 MB.
- The scraper runs **two Chrome instances in parallel** (one per portal).
- The output file is stored in the container's `/app` directory and served on demand.
  It resets on each redeploy — download it before redeploying.
