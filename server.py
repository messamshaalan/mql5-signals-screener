#!/usr/bin/env python3
"""
MT5 Signals Screener — Interactive Server
Usage: python server.py
Opens at http://localhost:5000  (Update button triggers live re-scrape)
"""

import os, sys, threading, time, webbrowser
from datetime import datetime
from flask import Flask, jsonify

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from scraper import (
    requests as cffi_requests,
    warm_up_session, page_url, fetch, parse_page,
    build_html, build_excel,
    enrich_with_symbols,
    TOTAL_PAGES, DELAY,
)
import pandas as pd

app   = Flask(__name__)
_OUT  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

_state = {
    "status":       "idle",
    "progress":     0,
    "total":        TOTAL_PAGES,
    "message":      "Ready",
    "last_updated": None,
    "html_path":    None,
    "signal_count": 0,
}
_lock = threading.Lock()


def _latest_html():
    if not os.path.exists(_OUT):
        return None
    files = sorted([f for f in os.listdir(_OUT) if f.endswith(".html")], reverse=True)
    return os.path.join(_OUT, files[0]) if files else None


@app.route("/")
def index():
    p = _state["html_path"] or _latest_html()
    if p and os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return f.read()
    return (
        '<html><body style="background:#0d1117;color:#c9d1d9;font-family:sans-serif;'
        'display:flex;align-items:center;justify-content:center;height:100vh;margin:0">'
        '<div style="text-align:center"><h2>No data yet</h2>'
        '<p>Click <strong>🔄 Update Data</strong> in the top-right to scrape MQL5.</p>'
        '<button onclick="fetch(\'/api/refresh\',{method:\'POST\'}).then(()=>location.reload())" '
        'style="background:#1F4E79;color:#fff;border:none;padding:10px 24px;border-radius:6px;'
        'font-size:1rem;cursor:pointer;margin-top:12px">🔄 Fetch Data Now</button>'
        '</div></body></html>'
    )


@app.route("/api/status")
def api_status():
    return jsonify(_state)


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    with _lock:
        if _state["status"] == "scraping":
            return jsonify({"error": "Already scraping"}), 409
        _state.update(status="scraping", progress=0, message="Starting…", signal_count=0)
    threading.Thread(target=_do_scrape, daemon=True).start()
    return jsonify({"started": True})


def _do_scrape():
    try:
        session     = cffi_requests.Session(impersonate="chrome120")
        all_signals = []

        with _lock:
            _state["message"] = "Warming up session…"
        warm_up_session(session)

        for page in range(1, TOTAL_PAGES + 1):
            url  = page_url(page)
            html = fetch(session, url)
            if html:
                sigs = parse_page(html, page)
                all_signals.extend(sigs)
            with _lock:
                _state.update(
                    progress=page,
                    message=f"Page {page}/{TOTAL_PAGES} — {len(all_signals)} signals",
                )
            if page < TOTAL_PAGES:
                time.sleep(DELAY)

        if not all_signals:
            with _lock:
                _state.update(status="error",
                              message="No signals scraped — MQL5 may be rate-limiting. Wait a few minutes and try again.")
            return

        # Apply symbols: name heuristic for new signals + cache for known ones (no detail-page fetch)
        all_signals = enrich_with_symbols(all_signals, session=None)

        df         = pd.DataFrame(all_signals)
        scraped_at = datetime.now().strftime("%Y-%m-%d %H:%M")
        ts         = datetime.now().strftime("%Y%m%d_%H%M")
        os.makedirs(_OUT, exist_ok=True)

        html_path  = os.path.join(_OUT, f"mt5_signals_{ts}.html")
        excel_path = os.path.join(_OUT, f"mt5_signals_{ts}.xlsx")

        build_html(df, html_path, scraped_at)
        build_excel(df, excel_path)

        with _lock:
            _state.update(
                status="done", html_path=html_path,
                last_updated=scraped_at, signal_count=len(all_signals),
                message=f"Done — {len(all_signals)} signals",
            )
    except Exception as e:
        with _lock:
            _state.update(status="error", message=str(e))


if __name__ == "__main__":
    os.makedirs(_OUT, exist_ok=True)
    print("=" * 55)
    print("  MT5 Signals Screener Server")
    print("  http://localhost:5000")
    print("=" * 55)
    threading.Timer(1.5, lambda: webbrowser.open("http://localhost:5000")).start()
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)
