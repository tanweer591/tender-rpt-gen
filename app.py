"""
app.py — Flask backend for the Defence Procurement Tender Scraper UI
=====================================================================
Endpoints
---------
GET  /          → serves tender_ui.html
POST /run       → starts BOTH scrapers in background threads (parallel)
GET  /status    → returns full dual-portal scraper state as JSON
GET  /download  → streams the finished .xlsx to the browser

Run
---
    pip install flask
    python app.py
"""

import os
import threading
from datetime import datetime
from flask import Flask, jsonify, send_file

from aerotender import scrape_all, write_excel, OUTPUT_FILE, KEYWORDS

app = Flask(__name__)

# ── Shared state ──────────────────────────────────────────────────────────────
# Written by the background thread(s); read by /status.
# Two sub-dicts — one per portal — so the UI can show parallel progress.

_state = {
    "status":         "idle",   # idle | running | done | error
    "total_keywords": len(KEYWORDS),
    "tenders_found":  0,
    "error":          "",
    "completed_at":   "",

    "defproc": {
        "status":          "idle",   # idle | running | done
        "keyword_index":   0,        # 0-based index of keyword currently being scanned
        "current_keyword": "",
        "found_indices":   [],       # which keyword indices returned results
        "tenders_found":   0,
    },

    "gem": {
        "status":          "idle",
        "keyword_index":   0,
        "current_keyword": "",
        "found_indices":   [],
        "tenders_found":   0,
    },
}

_lock = threading.Lock()


def _set(key, value):
    """Top-level state setter (thread-safe)."""
    with _lock:
        _state[key] = value


def _portal_set(portal, key, value):
    """Portal-specific state setter (thread-safe)."""
    with _lock:
        _state[portal][key] = value


# ── Background scraper ────────────────────────────────────────────────────────

def _run_scraper():
    """
    Runs in a daemon thread.
    Monkey-patches BOTH search_keyword (DefProc) AND scrape_gem_keyword (GeM)
    so each keyword search reports live progress back to _state.
    """
    # ── Reset state ──────────────────────────────────────────────────────────
    with _lock:
        _state["status"]        = "running"
        _state["tenders_found"] = 0
        _state["error"]         = ""
        _state["completed_at"]  = ""

        for portal in ("defproc", "gem"):
            _state[portal]["status"]          = "idle"
            _state[portal]["keyword_index"]   = 0
            _state[portal]["current_keyword"] = ""
            _state[portal]["found_indices"]   = []
            _state[portal]["tenders_found"]   = 0

    try:
        import aerotender as _at

        # ── Instrument DefProc ────────────────────────────────────────────────
        _orig_defproc = _at.search_keyword

        def _wrap_defproc(driver, keyword):
            idx = KEYWORDS.index(keyword) if keyword in KEYWORDS else -1
            with _lock:
                _state["defproc"]["status"]          = "running"
                _state["defproc"]["current_keyword"] = keyword
                _state["defproc"]["keyword_index"]   = idx

            rows = _orig_defproc(driver, keyword)

            if rows:
                with _lock:
                    _state["defproc"]["found_indices"].append(idx)

            # Advance index past the keyword just finished
            with _lock:
                _state["defproc"]["keyword_index"] = idx + 1

            return rows

        _at.search_keyword = _wrap_defproc

        # ── Instrument GeM ────────────────────────────────────────────────────
        _orig_gem = _at.scrape_gem_keyword

        def _wrap_gem(driver, keyword):
            idx = KEYWORDS.index(keyword) if keyword in KEYWORDS else -1
            with _lock:
                _state["gem"]["status"]          = "running"
                _state["gem"]["current_keyword"] = keyword
                _state["gem"]["keyword_index"]   = idx

            rows = _orig_gem(driver, keyword)

            if rows:
                with _lock:
                    _state["gem"]["found_indices"].append(idx)

            with _lock:
                _state["gem"]["keyword_index"] = idx + 1

            return rows

        _at.scrape_gem_keyword = _wrap_gem

        # ── Run both portals in parallel (mirrors aerotender.scrape_all) ─────
        tenders = scrape_all()

        # ── Restore originals ─────────────────────────────────────────────────
        _at.search_keyword    = _orig_defproc
        _at.scrape_gem_keyword = _orig_gem

        # ── Write Excel ───────────────────────────────────────────────────────
        if tenders:
            write_excel(tenders)

        defproc_count = sum(1 for t in tenders if t.get("Source") == "DefProc")
        gem_count     = sum(1 for t in tenders if t.get("Source") == "GeM")

        with _lock:
            _state["status"]        = "done"
            _state["tenders_found"] = len(tenders)
            _state["completed_at"]  = datetime.now().strftime("%H:%M")

            _state["defproc"]["status"]        = "done"
            _state["defproc"]["tenders_found"] = defproc_count
            _state["defproc"]["keyword_index"] = len(KEYWORDS)

            _state["gem"]["status"]        = "done"
            _state["gem"]["tenders_found"] = gem_count
            _state["gem"]["keyword_index"] = len(KEYWORDS)

    except Exception as exc:
        with _lock:
            _state["status"] = "error"
            _state["error"]  = str(exc)


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    with open("tender_ui.html", encoding="utf-8") as f:
        return f.read()


@app.route("/run", methods=["POST"])
def run():
    with _lock:
        if _state["status"] == "running":
            return jsonify({"ok": False, "reason": "already running"}), 409

    t = threading.Thread(target=_run_scraper, daemon=True)
    t.start()
    return jsonify({"ok": True})


@app.route("/status")
def status():
    with _lock:
        return jsonify(dict(
            status         = _state["status"],
            total_keywords = _state["total_keywords"],
            tenders_found  = _state["tenders_found"],
            error          = _state["error"],
            completed_at   = _state["completed_at"],
            defproc        = dict(_state["defproc"]),
            gem            = dict(_state["gem"]),
        ))


@app.route("/download")
def download():
    if not os.path.exists(OUTPUT_FILE):
        return jsonify({"error": "File not ready"}), 404
    return send_file(
        OUTPUT_FILE,
        as_attachment=True,
        download_name="tenders_output.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
    )


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("\n  Defence Tender Scraper UI")
    print("  Open http://localhost:5000 in your browser\n")
    app.run(debug=False, port=5000)
