"""
Placement Offers — local web dashboard
======================================
One-click UI over the Telegram scanner. Run this (or double-click the .bat),
your browser opens, you Scan, browse offers in clean boxes with AIML / BACKEND /
ALL tabs, search + filters, and download a tidy Excel file. No terminal needed.
"""

import json
import os
import subprocess
import sys
import threading
import webbrowser

from flask import Flask, jsonify, request, send_file, Response

import placement_monitor as pm
import db

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
PORT = 5005

# --- Durable backend: SQLite, with one-time import of any legacy JSON files ---
db.init()
try:
    _legacy_applied = None
    _ap = os.path.join(HERE, "applied.json")
    if os.path.exists(_ap):
        with open(_ap, encoding="utf-8") as _f:
            _legacy_applied = json.load(_f)
    _legacy_seen = None
    _sp = os.path.join(HERE, "seen.json")
    if os.path.exists(_sp):
        with open(_sp, encoding="utf-8") as _f:
            _legacy_seen = int(json.load(_f).get("last_seen_id", 0))
    db.migrate_from_json(_legacy_applied, _legacy_seen)
except Exception as _e:  # noqa
    print("migration note:", _e)

app = Flask(__name__)
_status = {"running": False, "message": "", "added": None, "error": None}


@app.after_request
def _no_cache(resp):
    # Never let the browser serve a stale page/data — you always get the latest.
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


def _run_scan(full):
    _status.update(running=True, message="Scanning Telegram…", added=None, error=None)
    try:
        args = [sys.executable, os.path.join(HERE, "placement_monitor.py"),
                "--full" if full else "--once"]
        proc = subprocess.run(args, cwd=HERE, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=1800)
        if proc.returncode != 0:
            _status.update(error=(proc.stderr or proc.stdout or "scan failed").strip()[-500:])
        else:
            out = (proc.stdout or "").strip()
            added = None
            for tok in out.replace("]", " ").split():
                if tok.isdigit():
                    added = int(tok)
                    break
            _status.update(added=added, message=out.splitlines()[-1] if out else "Done")
    except subprocess.TimeoutExpired:
        _status.update(error="Scan timed out (very large group). Try again.")
    except Exception as e:  # noqa
        _status.update(error=str(e))
    finally:
        _status["running"] = False


@app.route("/")
def index():
    return Response(INDEX_HTML, mimetype="text/html")


@app.route("/api/offers")
def api_offers():
    offers = pm.grouped_offers()
    applied = db.get_applied()
    seen_offers = db.get_seen_offers()
    for o in offers:
        rec = applied.get(o["oid"])
        o["applied"] = bool(rec and rec.get("applied"))
        o["applied_date"] = (rec or {}).get("date", "")
    # "New" = an offer that arrived after the baseline AND hasn't been dismissed.
    global_max = max((o.get("last_id", 0) for o in offers), default=0)
    baseline = db.get_meta("last_seen_id")
    if baseline is None:            # first ever load: nothing is new yet
        db.set_meta("last_seen_id", global_max)
        baseline = global_max
    baseline = int(baseline)
    for o in offers:
        o["is_new"] = (o.get("last_id", 0) > baseline) and (o["oid"] not in seen_offers)
    counts = {
        "ALL": len(offers),
        "AIML": sum("AIML" in o["categories"] for o in offers),
        "BACKEND": sum("BACKEND" in o["categories"] for o in offers),
        "with_link": sum(o["has_link"] for o in offers),
        "open": sum(o["status"] == "OPEN" for o in offers),
        "JOB": sum("JOB" in o["types"] for o in offers),
        "INTERN": sum("INTERN" in o["types"] for o in offers),
        "HACK": sum(o.get("kind") == "Hackathon" for o in offers),
        "applied": sum(o["applied"] for o in offers),
        "new": sum(o["is_new"] for o in offers),
    }
    return jsonify({"offers": offers, "counts": counts, "group": pm.GROUP,
                    "series": db.applied_series()})


@app.route("/api/seen_offer", methods=["POST"])
def api_seen_offer():
    """Dismiss the NEW tag on one offer (the per-card tick)."""
    body = request.get_json(force=True, silent=True) or {}
    oid = str(body.get("oid", "")).strip()
    if not oid:
        return jsonify({"ok": False, "error": "missing oid"}), 400
    db.mark_seen_offer(oid, bool(body.get("seen", True)))
    return jsonify({"ok": True})


@app.route("/api/seen", methods=["POST"])
def api_seen():
    """Dismiss NEW on all currently-new offers at once."""
    offers = pm.grouped_offers()
    seen_offers = db.get_seen_offers()
    baseline = int(db.get_meta("last_seen_id") or 0)
    new_oids = [o["oid"] for o in offers
                if o.get("last_id", 0) > baseline and o["oid"] not in seen_offers]
    db.mark_seen_offers(new_oids)
    return jsonify({"ok": True, "new": 0})


@app.route("/api/applied", methods=["POST"])
def api_applied():
    body = request.get_json(force=True, silent=True) or {}
    oid = str(body.get("oid", "")).strip()
    if not oid:
        return jsonify({"ok": False, "error": "missing oid"}), 400
    want = bool(body.get("applied"))
    total = db.set_applied(oid, want)
    return jsonify({"ok": True, "applied": want, "total": total,
                    "series": db.applied_series()})


@app.route("/api/scan", methods=["POST"])
def api_scan():
    if _status["running"]:
        return jsonify({"ok": False, "error": "A scan is already running."}), 409
    full = request.args.get("mode") == "full"
    threading.Thread(target=_run_scan, args=(full,), daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/status")
def api_status():
    return jsonify(_status)


@app.route("/download")
def download():
    category = request.args.get("category", "ALL").upper()
    if category not in ("ALL", "AIML", "BACKEND"):
        category = "ALL"
    out = os.path.join(HERE, f"placement_offers_{category.lower()}.xlsx")
    n = pm.build_excel(path=out, category=category)
    if n == 0:
        return "No offers yet — run a scan first.", 404
    return send_file(out, as_attachment=True,
                     download_name=f"placement_offers_{category.lower()}.xlsx")


# --------------------------------------------------------------------------- #
INDEX_HTML = r"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Placement Offers</title>
<style>
  :root{
    /* Neutral + accent are overridden by the theme switcher (JS). */
    --bg:#f1f5f9; --panel:#ffffff; --line:#e2e8f0; --text:#1e293b; --muted:#64748b;
    --accent:#4f46e5; --accent-weak:#e8e7fb; --chip:#eef2f7;
    --shadow:0 1px 2px rgba(15,23,42,.06),0 4px 14px rgba(15,23,42,.06);
    /* Semantic colors (shared across themes). */
    --green:#0a8f5b; --green-weak:#e2f6ec; --purple:#7b3fe4; --purple-weak:#f0e9fd;
    --amber:#a86400; --amber-weak:#fdf1dd; --pink:#c026a3; --pink-weak:#fbe6f6;
  }
  @media (prefers-color-scheme: dark){
    :root{
      --bg:#0b1120; --panel:#151c2c; --line:#232c40; --text:#e6ebf5; --muted:#94a3b8;
      --accent:#818cf8; --accent-weak:#1e2340; --chip:#1c2536;
      --shadow:0 1px 2px rgba(0,0,0,.3),0 8px 20px rgba(0,0,0,.3);
      --green:#37d39a; --green-weak:#123027; --purple:#b18bff; --purple-weak:#241a3a;
      --amber:#e0a44a; --amber-weak:#2e2413; --pink:#f19adb; --pink-weak:#331a2c;
    }
  }
  *{box-sizing:border-box}
  html,body{margin:0}
  body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--bg);
    color:var(--text);font-size:14.5px;line-height:1.5}
  .top{position:sticky;top:0;z-index:20;background:var(--panel);
    border-bottom:1px solid var(--line);box-shadow:var(--shadow)}
  .top-in{max-width:1120px;margin:0 auto;padding:14px 22px;display:flex;
    align-items:center;gap:14px;flex-wrap:wrap}
  .brand{font-size:18px;font-weight:800;letter-spacing:-.2px;display:flex;gap:8px;align-items:center}
  .brand small{font-weight:500;color:var(--muted);font-size:12px}
  .grow{flex:1}
  .btn{font:inherit;font-weight:650;border:1px solid transparent;border-radius:9px;
    padding:9px 15px;cursor:pointer;color:#fff;background:var(--accent);transition:.15s;
    display:inline-flex;gap:7px;align-items:center}
  .btn:hover{filter:brightness(1.06)} .btn:disabled{opacity:.55;cursor:not-allowed}
  .btn.sec{background:transparent;color:var(--text);border-color:var(--line)}
  .btn.sec:hover{background:var(--chip)}
  .btn.grn{background:var(--green)}
  .palette-sel{font:inherit;font-weight:600;font-size:13px;background:var(--chip);
    color:var(--text);border:1px solid var(--line);border-radius:9px;padding:9px 11px;cursor:pointer}
  .wrap{max-width:1120px;margin:0 auto;padding:20px 22px 70px}
  .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:18px}
  .stat{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;box-shadow:var(--shadow)}
  .stat .n{font-size:24px;font-weight:800;letter-spacing:-.5px}
  .stat .l{color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.05em;margin-top:3px;font-weight:600}
  .toolbar{position:sticky;top:60px;z-index:15;display:flex;gap:10px;align-items:center;
    flex-wrap:wrap;background:var(--bg);padding:8px 0 14px}
  .tabs{display:inline-flex;background:var(--panel);border:1px solid var(--line);
    border-radius:10px;padding:4px;gap:2px;box-shadow:var(--shadow)}
  .tab{border:0;background:transparent;font:inherit;font-weight:700;color:var(--muted);
    padding:8px 16px;border-radius:7px;cursor:pointer;transition:.12s;white-space:nowrap}
  .tab .c{font-weight:600;opacity:.75;margin-left:5px;font-size:12px}
  .tab.on{background:var(--accent);color:#fff}
  .tab.on .c{opacity:.9}
  .search{flex:1;min-width:180px;position:relative}
  .search input{width:100%;font:inherit;padding:10px 12px 10px 34px;border-radius:9px;
    border:1px solid var(--line);background:var(--panel);color:var(--text);box-shadow:var(--shadow)}
  .search::before{content:'🔎';position:absolute;left:11px;top:9px;opacity:.6}
  .toggle{display:inline-flex;gap:8px;align-items:center;color:var(--muted);
    font-weight:600;font-size:13px;user-select:none;cursor:pointer;
    background:var(--panel);border:1px solid var(--line);padding:9px 13px;border-radius:9px;box-shadow:var(--shadow)}
  .toggle input{accent-color:var(--accent);width:16px;height:16px}
  .toggle.apply-now{border-color:var(--green);color:var(--green);font-weight:700}
  .toggle.apply-now input{accent-color:var(--green)}
  .cgpa{display:inline-flex;align-items:center;gap:9px;background:var(--panel);
    border:1px solid var(--line);padding:8px 13px;border-radius:9px;box-shadow:var(--shadow);
    font-weight:600;font-size:13px;color:var(--muted)}
  .cgpa input[type=range]{accent-color:var(--accent);width:120px;cursor:pointer}
  .cgpa b{color:var(--accent);font-size:14px;min-width:30px;text-align:center}
  .cgpa-tag{background:var(--accent-weak);color:var(--accent);font-weight:700;
    padding:2px 9px;border-radius:6px;font-size:13px}
  .banner{padding:11px 15px;border-radius:11px;margin-bottom:14px;font-size:13.5px;
    display:none;border:1px solid var(--line);background:var(--panel);box-shadow:var(--shadow)}
  .banner.show{display:block}
  .banner.err{background:#fdeaea;border-color:#f5c2c2;color:#9c2b2b}
  @media (prefers-color-scheme:dark){.banner.err{background:#2c1618;border-color:#5a2b2e;color:#ffb4b4}}
  .newbar{display:none;align-items:center;gap:12px;background:var(--accent-weak);
    border:1px solid var(--accent);border-radius:12px;padding:11px 16px;margin-bottom:14px;
    font-weight:700;font-size:14px;color:var(--accent)}
  .newbar.show{display:flex}
  .newbar .grow{flex:1}
  .newbar button{font:inherit;font-weight:700;font-size:13px;border:1px solid var(--accent);
    background:transparent;color:var(--accent);padding:7px 13px;border-radius:8px;cursor:pointer}
  .newbar button:hover{background:var(--accent);color:#fff}
  .tag.NEW{background:var(--accent);color:#fff;animation:newpulse 1.8s ease-in-out infinite}
  @keyframes newpulse{0%,100%{opacity:1}50%{opacity:.55}}
  .newtag{display:inline-flex;align-items:center;gap:5px;background:var(--accent);color:#fff;
    font:inherit;font-size:10px;font-weight:800;padding:4px 9px;border:0;border-radius:999px;cursor:pointer;
    text-transform:uppercase;letter-spacing:.03em;user-select:none;animation:newpulse 1.8s ease-in-out infinite}
  .newtag:hover{filter:brightness(1.12);animation:none}
  .card.new{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent),var(--shadow)}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:16px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:16px;
    padding:18px 19px 16px;box-shadow:var(--shadow);display:flex;flex-direction:column;
    transition:transform .16s ease,box-shadow .16s ease,border-color .16s ease;position:relative}
  .card:hover{transform:translateY(-3px);border-color:var(--accent);
    box-shadow:0 2px 6px rgba(20,30,50,.08),0 12px 28px rgba(20,30,50,.12)}
  .card .top-row{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:2px}
  .company{font-size:16.5px;font-weight:800;letter-spacing:-.2px;margin:0;line-height:1.25;
    display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .chips{display:flex;flex-flow:row wrap;gap:5px;justify-content:flex-end;align-items:center;max-width:52%;flex-shrink:0}
  .tag{font-size:10px;font-weight:800;padding:3px 8px;border-radius:999px;white-space:nowrap;text-transform:uppercase;letter-spacing:.03em}
  .tag.JOB{background:var(--accent-weak);color:var(--accent)}
  .tag.INTERN{background:var(--amber-weak);color:var(--amber)}
  .tag.AIML{background:var(--purple-weak);color:var(--purple)}
  .tag.BACKEND{background:var(--green-weak);color:var(--green)}
  .tag.General{background:var(--chip);color:var(--muted)}
  .tag.HACK{background:var(--pink-weak);color:var(--pink)}
  .sub{color:var(--text);opacity:.72;font-size:13px;font-weight:600;margin:6px 0 11px;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .sec{margin:0 0 11px}
  .mini{font-size:9.5px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;
    color:var(--muted);display:block;margin-bottom:4px}
  .sec .val{font-size:12.5px;color:var(--text);opacity:.9;line-height:1.45;
    display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .pills{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:12px}
  .pill{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;font-weight:700;
    padding:4px 10px;border-radius:7px;background:var(--chip);color:var(--muted);white-space:nowrap}
  .pill.neutral{background:var(--accent-weak);color:var(--accent)}
  .pill.OPEN{background:var(--green-weak);color:var(--green)}
  .pill.CLOSED{background:#f2ddd3;color:#a4472e}
  .pill.UNKNOWN{background:var(--chip);color:var(--muted)}
  .pill.URGENT{background:var(--amber-weak);color:var(--amber);
    box-shadow:0 0 0 1px var(--amber) inset;font-weight:800}
  @media (prefers-color-scheme:dark){.pill.CLOSED{background:#3a211a;color:#e8a892}}
  .skills{display:flex;flex-wrap:wrap;gap:6px}
  .skill{background:var(--accent-weak);color:var(--accent);font-size:11px;font-weight:700;
    padding:3px 9px;border-radius:6px;border:1px solid transparent}
  .foot{margin-top:auto;padding-top:14px;border-top:1px solid var(--line);
    display:flex;align-items:center;gap:8px}
  .apply{flex:1;text-align:center;background:var(--accent);color:#fff;text-decoration:none;
    font-weight:700;font-size:13px;padding:10px;border-radius:9px;transition:filter .15s}
  .apply:hover{filter:brightness(1.08)}
  .nolink{flex:1;text-align:center;background:var(--amber-weak);color:var(--amber);
    font-weight:700;font-size:12.5px;padding:9px;border-radius:8px}
  .times{font-size:11.5px;color:var(--muted);font-weight:600;white-space:nowrap}
  .empty{grid-column:1/-1;text-align:center;color:var(--muted);padding:60px 20px}
  .spin{width:14px;height:14px;border:2px solid #fff6;border-top-color:#fff;border-radius:50%;
    display:inline-block;animation:sp .7s linear infinite;vertical-align:-2px}
  @keyframes sp{to{transform:rotate(360deg)}}
  /* Applications chart */
  .chartcard{background:var(--panel);border:1px solid var(--line);border-radius:14px;
    padding:16px 18px;box-shadow:var(--shadow);margin-bottom:18px}
  .chart-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
  .chart-title{font-weight:800;font-size:14.5px}
  .chart-total{font-size:12.5px;color:var(--muted);font-weight:700;
    background:var(--accent-weak);color:var(--accent);padding:3px 10px;border-radius:999px}
  .bars{display:flex;align-items:stretch;gap:16px;height:150px;overflow-x:auto;padding-top:4px}
  .bar-col{display:flex;flex-direction:column;align-items:center;min-width:42px;flex:1;height:100%}
  .bar-val{font-size:12px;font-weight:800;color:var(--text);margin-bottom:5px}
  .bar-track{flex:1;width:100%;display:flex;align-items:flex-end;justify-content:center}
  .bar{width:62%;max-width:46px;min-height:4px;background:var(--accent);
    border-radius:7px 7px 0 0;transition:height .35s ease;opacity:.92}
  .bar-col:hover .bar{opacity:1}
  .bar-day{font-size:10.5px;color:var(--muted);margin-top:7px;font-weight:600;white-space:nowrap}
  .chart-empty{color:var(--muted);font-size:13px;text-align:center;padding:14px 8px;line-height:1.5}
  /* Applied state */
  .applied-btn{background:transparent;border:1px solid var(--line);color:var(--muted);
    font:inherit;font-weight:700;font-size:12px;padding:9px 12px;border-radius:9px;cursor:pointer;
    display:inline-flex;gap:6px;align-items:center;transition:.15s;white-space:nowrap}
  .applied-btn:hover{border-color:var(--green);color:var(--green)}
  .applied-btn.on{background:var(--green-weak);border-color:var(--green);color:var(--green)}
  .card.applied{border-color:var(--green)}
  .card.applied::after{content:'✓ APPLIED';position:absolute;top:12px;right:14px;
    font-size:9px;font-weight:800;letter-spacing:.06em;color:var(--green);opacity:.0}
  .ribbon{position:absolute;top:0;left:0;width:100%;height:3px;background:var(--green);
    border-radius:16px 16px 0 0;display:none}
  .card.applied .ribbon{display:block}
</style></head>
<body>
<div class="top"><div class="top-in">
  <div class="brand">🎯 Placement Offers <small id="grouplabel"></small></div>
  <div class="grow"></div>
  <select id="palette" class="palette-sel" onchange="setPalette(this.value)" title="Color theme">
    <option value="indigo">🎨 Indigo Slate</option>
    <option value="emerald">🎨 Emerald Noir</option>
    <option value="violet">🎨 Royal Violet</option>
    <option value="nord">🎨 Nord Frost</option>
  </select>
  <button class="btn" id="scanBtn" onclick="scan('new')">🔄 Scan new</button>
  <button class="btn sec" onclick="scan('full')">↻ Full rescan</button>
  <button class="btn grn" onclick="dl()">⬇ Download Excel</button>
</div></div>

<div class="wrap">
  <div id="banner" class="banner"></div>
  <div id="newbar" class="newbar"></div>
  <div class="stats" id="stats"></div>

  <div class="toolbar">
    <div class="tabs" id="tabs">
      <button class="tab on" data-cat="ALL" onclick="setCat('ALL')">All <span class="c" id="c-ALL"></span></button>
      <button class="tab" data-cat="AIML" onclick="setCat('AIML')">AI / ML <span class="c" id="c-AIML"></span></button>
      <button class="tab" data-cat="BACKEND" onclick="setCat('BACKEND')">Backend <span class="c" id="c-BACKEND"></span></button>
    </div>
    <div class="tabs" id="typeTabs">
      <button class="tab on" data-type="ANY" onclick="setType('ANY')">All</button>
      <button class="tab" data-type="JOB" onclick="setType('JOB')">Jobs <span class="c" id="t-JOB"></span></button>
      <button class="tab" data-type="INTERN" onclick="setType('INTERN')">Internships <span class="c" id="t-INTERN"></span></button>
      <button class="tab" data-type="HACK" onclick="setType('HACK')">Hackathons <span class="c" id="t-HACK"></span></button>
    </div>
    <label class="toggle apply-now"><input type="checkbox" id="applyNow" onchange="render()"> ⚡ Apply now</label>
    <label class="toggle"><input type="checkbox" id="newOnly" onchange="render()"> 🆕 New only</label>
    <label class="toggle"><input type="checkbox" id="linkOnly" onchange="render()"> With link only</label>
    <select id="appFilter" class="palette-sel" onchange="render()" title="Applied status">
      <option value="all">All offers</option>
      <option value="todo">To-apply</option>
      <option value="done">✓ Applied</option>
    </select>
    <div class="cgpa">
      <span>My CGPA</span>
      <input type="range" id="cgpa" min="4" max="10" step="0.1" value="10" oninput="render()">
      <b id="cgpaVal">10.0</b>
    </div>
    <div class="search"><input id="q" placeholder="Search company or skill…" oninput="render()"></div>
  </div>

  <div class="grid" id="grid"></div>
</div>

<script>
let DATA={offers:[],counts:{}}, CAT='ALL', TYPE='ANY';

// ---- Color themes (neutral + accent swapped live; semantic colors stay) ----
const PALETTES={
  indigo:{
    light:{bg:'#f1f5f9',panel:'#ffffff',line:'#e2e8f0',text:'#1e293b',muted:'#64748b',accent:'#4f46e5','accent-weak':'#e8e7fb',chip:'#eef2f7',shadow:'0 1px 2px rgba(15,23,42,.06),0 4px 14px rgba(15,23,42,.06)'},
    dark:{bg:'#0b1120',panel:'#151c2c',line:'#232c40',text:'#e6ebf5',muted:'#94a3b8',accent:'#818cf8','accent-weak':'#1e2340',chip:'#1c2536',shadow:'0 1px 2px rgba(0,0,0,.3),0 8px 20px rgba(0,0,0,.3)'}},
  emerald:{
    light:{bg:'#eef4f0',panel:'#ffffff',line:'#d9e7df',text:'#13211a',muted:'#5c766a',accent:'#059669','accent-weak':'#d6f3e6',chip:'#e6f1eb',shadow:'0 1px 2px rgba(6,40,30,.06),0 4px 14px rgba(6,40,30,.07)'},
    dark:{bg:'#0f1512',panel:'#18201c',line:'#26332c',text:'#e7f0ea',muted:'#93a89d',accent:'#10b981','accent-weak':'#123026',chip:'#1a231e',shadow:'0 1px 2px rgba(0,0,0,.35),0 8px 20px rgba(0,0,0,.32)'}},
  violet:{
    light:{bg:'#f4f3fb',panel:'#ffffff',line:'#e7e3f4',text:'#211f3a',muted:'#6d6a8c',accent:'#7c3aed','accent-weak':'#ece7fd',chip:'#efedf8',shadow:'0 1px 2px rgba(30,25,60,.06),0 4px 14px rgba(30,25,60,.07)'},
    dark:{bg:'#12132a',panel:'#1c1d3a',line:'#2c2d4e',text:'#e8e7f6',muted:'#a3a1c4',accent:'#8b5cf6','accent-weak':'#282a52',chip:'#20213f',shadow:'0 1px 2px rgba(0,0,0,.35),0 8px 22px rgba(0,0,0,.34)'}},
  nord:{
    light:{bg:'#eceff4',panel:'#ffffff',line:'#dde3ec',text:'#2e3440',muted:'#5b6472',accent:'#5e81ac','accent-weak':'#e2ebf3',chip:'#e7ecf3',shadow:'0 1px 2px rgba(46,52,64,.06),0 4px 14px rgba(46,52,64,.07)'},
    dark:{bg:'#242933',panel:'#2e3440',line:'#3b4252',text:'#e5e9f0',muted:'#a7b1c2',accent:'#88c0d0','accent-weak':'#3a4457',chip:'#333a49',shadow:'0 1px 2px rgba(0,0,0,.35),0 8px 20px rgba(0,0,0,.3)'}},
};
function isDark(){return window.matchMedia&&window.matchMedia('(prefers-color-scheme: dark)').matches;}
function setPalette(name){
  const p=(PALETTES[name]||PALETTES.indigo)[isDark()?'dark':'light'];
  for(const k in p) document.documentElement.style.setProperty('--'+k, p[k]);
  localStorage.setItem('ppal', name);
  const sel=document.getElementById('palette'); if(sel) sel.value=name;
}
(function initTheme(){
  const saved=localStorage.getItem('ppal')||'indigo';
  setPalette(saved);
  if(window.matchMedia) window.matchMedia('(prefers-color-scheme: dark)')
    .addEventListener('change',()=>setPalette(localStorage.getItem('ppal')||'indigo'));
})();

function esc(s){return (s==null?'':''+s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function banner(msg,kind){const b=document.getElementById('banner');
  if(!msg){b.className='banner';return;} b.className='banner show '+(kind||''); b.innerHTML=msg;}

// Per-browser dismissed-NEW overlay — makes a dismissal immune to any re-fetch
// race (server stays the durable source via seen_offers).
let DISMISSED=new Set();
try{ DISMISSED=new Set(JSON.parse(localStorage.getItem('dismissed_oids')||'[]')); }catch(e){}
function saveDismissed(){ try{ localStorage.setItem('dismissed_oids',JSON.stringify([...DISMISSED])); }catch(e){} }
function applyDismissed(offers){
  if(!offers) return;
  for(const o of offers){
    if(!DISMISSED.has(o.oid)) continue;
    if(o.is_new) o.is_new=false;        // server not caught up yet -> keep hidden
    else DISMISSED.delete(o.oid);       // server already agrees -> stop tracking (no drift)
  }
  saveDismissed();
}

async function load(){
  const r=await fetch('/api/offers'); DATA=await r.json();
  applyDismissed(DATA.offers);
  DATA.counts.new=DATA.offers.filter(o=>o.is_new).length;
  document.getElementById('grouplabel').textContent='· group '+(DATA.group||'');
  for(const k of ['ALL','AIML','BACKEND']) document.getElementById('c-'+k).textContent=DATA.counts[k]??'';
  document.getElementById('t-JOB').textContent=DATA.counts.JOB??'';
  document.getElementById('t-INTERN').textContent=DATA.counts.INTERN??'';
  document.getElementById('t-HACK').textContent=DATA.counts.HACK??'';
  renderStats(); renderChart(DATA.series); renderNewbar(); render();
}
function renderStats(){
  const c=DATA.counts||{};
  const cards=[['Total offers',c.ALL||0],['New',c.new||0],
    ['Applied',c.applied||0],['Open now',c.open||0]];
  document.getElementById('stats').innerHTML=cards.map(
    x=>`<div class="stat"><div class="n">${x[1]}</div><div class="l">${x[0]}</div></div>`).join('');
}
function renderNewbar(){
  const n=DATA.counts?.new||0;
  const bar=document.getElementById('newbar');
  if(!n){bar.className='newbar';bar.innerHTML='';return;}
  bar.className='newbar show';
  bar.innerHTML=`<span>🆕 ${n} new offer${n>1?'s':''} since your last check</span>`
    +`<span class="grow"></span>`
    +`<button onclick="filterNew()">Show new</button>`
    +`<button onclick="markSeen()">Mark all seen</button>`;
}
function filterNew(){document.getElementById('newOnly').checked=true;render();
  window.scrollTo({top:document.querySelector('.toolbar').offsetTop-70,behavior:'smooth'});}
async function dismissNew(e, oid){
  e.stopPropagation(); e.preventDefault();
  DISMISSED.add(oid); saveDismissed();            // sticky locally (race-proof)
  const o=DATA.offers.find(x=>x.oid===oid); if(o){o.is_new=false;}
  DATA.counts.new=DATA.offers.filter(x=>x.is_new).length;
  renderStats(); renderNewbar(); render();
  try{ await fetch('/api/seen_offer',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({oid,seen:true})}); }catch(err){}                // durable on server
}
async function markSeen(){
  DATA.offers.forEach(o=>{ if(o.is_new){DISMISSED.add(o.oid);} o.is_new=false; });
  saveDismissed(); DATA.counts.new=0;
  document.getElementById('newOnly').checked=false;
  renderStats(); renderNewbar(); render();
  try{ await fetch('/api/seen',{method:'POST'}); }catch(e){}
}
function renderChart(series){
  const el=document.getElementById('chartcard');
  if(!el) return;                      // graph removed — no-op
  if(!series||!series.length){
    el.innerHTML=`<div class="chart-empty">📊 No applications logged yet — tick <b>☐ Applied</b> on offers you apply to, and your daily progress builds up here.</div>`;
    return;}
  const total=series.reduce((a,s)=>a+s.count,0);
  const max=Math.max(...series.map(s=>s.count),1);
  const bars=series.map(s=>{
    const h=Math.max(Math.round(s.count/max*100),8);
    return `<div class="bar-col" title="${s.date}: ${s.count} applied">`+
      `<div class="bar-val">${s.count}</div>`+
      `<div class="bar-track"><div class="bar" style="height:${h}%"></div></div>`+
      `<div class="bar-day">${s.date.slice(5)}</div></div>`;
  }).join('');
  el.innerHTML=`<div class="chart-head"><span class="chart-title">📊 Applications over time</span><span class="chart-total">${total} total</span></div>`+
    `<div class="bars">${bars}</div>`;
}
function setCat(c){CAT=c;
  document.querySelectorAll('#tabs .tab').forEach(t=>t.classList.toggle('on',t.dataset.cat===c));
  render();}
function setType(t){TYPE=t;
  document.querySelectorAll('#typeTabs .tab').forEach(x=>x.classList.toggle('on',x.dataset.type===t));
  render();}
function render(){
  const q=document.getElementById('q').value.trim().toLowerCase();
  const linkOnly=document.getElementById('linkOnly').checked;
  const applyNow=document.getElementById('applyNow').checked;
  const newOnly=document.getElementById('newOnly').checked;
  const appFilter=document.getElementById('appFilter').value;
  const myCgpa=parseFloat(document.getElementById('cgpa').value);
  document.getElementById('cgpaVal').textContent=myCgpa.toFixed(1);
  let rows=DATA.offers.filter(o=>{
    if(CAT!=='ALL' && !o.categories.includes(CAT)) return false;
    if(TYPE==='HACK'){ if(o.kind!=='Hackathon') return false; }
    else if(TYPE!=='ANY' && !o.types.includes(TYPE)) return false;
    if(newOnly && !o.is_new) return false;
    if(appFilter==='done' && !o.applied) return false;
    if(appFilter==='todo' && o.applied) return false;
    if(applyNow && !(o.status==='OPEN' && o.has_link)) return false;
    if(linkOnly && !o.has_link) return false;
    // CGPA cutoff: drop offers that need MORE than the user's CGPA.
    // Offers with no stated CGPA always stay.
    if(o.cgpa!=null && o.cgpa>myCgpa) return false;
    if(q){const hay=(o.Company+' '+(o.skills||[]).join(' ')+' '+(o['Role / Position']||'')).toLowerCase();
      if(!hay.includes(q)) return false;}
    return true;
  });
  rows.sort((a,b)=>(b.is_new?1:0)-(a.is_new?1:0));  // new offers first (stable)
  const grid=document.getElementById('grid');
  if(!rows.length){grid.innerHTML='<div class="empty">No offers match.<br>Try another tab, clear search, or click <b>Scan new</b>.</div>';return;}
  grid.innerHTML=rows.map(card).join('');
}
const STATUS_LABEL={OPEN:'Open',CLOSED:'Closed',UNKNOWN:'Deadline N/A'};
function daysLeft(o){
  if(o.status!=='OPEN'||o.days_left==null) return '';
  const d=o.days_left;
  const txt=d<=0?'last day!':(d===1?'1 day left':d+' days left');
  const urgent=d<=2?' urgent':'';
  return ` <span class="left${urgent}">(${txt})</span>`;
}
function daysLeftShort(o){
  if(o.status!=='OPEN'||o.days_left==null) return '';
  const d=o.days_left;
  return ' · <b>'+(d<=0?'last day!':(d===1?'1 day left':d+' days left'))+'</b>';
}
function card(o){
  const cat=o.primary||'General';
  const st=o.status||'UNKNOWN';
  const isHack=o.kind==='Hackathon';
  const typeCls=(o.types.includes('INTERN')&&!o.types.includes('JOB'))?'INTERN':'JOB';
  const chips=(o.is_new?`<button class="newtag" title="Click to dismiss the New tag" onclick="dismissNew(event,'${esc(o.oid)}')">🆕 New ✕</button>`:'')
    +(isHack?`<span class="tag HACK">Hackathon</span>`:'')
    +`<span class="tag ${typeCls}">${esc(o.type_label)}</span>`
    +`<span class="tag ${cat}">${cat==='General'?'General':cat}</span>`;
  const sub=o.summary?`<div class="sub" title="${esc(o.summary)}">${esc(o.summary)}</div>`:'';
  const elig=o['Eligibility / Batch']
    ? `<div class="sec"><span class="mini">Eligibility</span><div class="val">${esc(o['Eligibility / Batch'])}</div></div>`:'';
  const pills=[];
  if(o.cgpa!=null) pills.push(`<span class="pill neutral">CGPA ≥ ${o.cgpa.toFixed(1)}</span>`);
  const urgent=(st==='OPEN' && o.days_left!=null && o.days_left<=2);
  const dlIcon=urgent?'⏰':'📅';
  pills.push(o.deadline
    ? `<span class="pill ${urgent?'URGENT':st}">${dlIcon} ${esc(o.deadline)} · ${STATUS_LABEL[st]}${daysLeftShort(o)}</span>`
    : `<span class="pill UNKNOWN">No deadline listed</span>`);
  const chipsSk=(o.skills||[]).slice(0,10).map(s=>`<span class="skill">${esc(s)}</span>`).join('');
  const skillsSec=chipsSk
    ? `<div class="sec"><span class="mini">Skills required</span><div class="skills">${chipsSk}</div></div>`:'';
  const foot=o.has_link
    ? `<a class="apply" href="${esc(o['Apply Link'])}" target="_blank" rel="noopener">Apply / Details ↗</a>`
    : `<span class="nolink">No link — check group</span>`;
  const appBtn=`<button class="applied-btn ${o.applied?'on':''}" title="${o.applied?'Applied on '+esc(o.applied_date):'Mark as applied'}" onclick="toggleApplied(event,'${esc(o.oid)}')">${o.applied?'✓ Applied':'☐ Applied'}</button>`;
  return `<div class="card${o.applied?' applied':''}${o.is_new?' new':''}" data-oid="${esc(o.oid)}">
    <div class="ribbon"></div>
    <div class="top-row">
      <h3 class="company">${esc(o.Company)}</h3>
      <div class="chips">${chips}</div>
    </div>
    ${sub}
    <div class="pills">${pills.join('')}</div>
    ${elig}${skillsSec}
    <div class="foot">${foot}${appBtn}</div>
  </div>`;
}
async function toggleApplied(e, oid){
  e.stopPropagation(); e.preventDefault();
  const o=DATA.offers.find(x=>x.oid===oid); if(!o) return;
  const next=!o.applied;
  o.applied=next; o.applied_date=next?(new Date().toISOString().slice(0,10)):'';
  render(); renderStats();
  try{
    const r=await fetch('/api/applied',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({oid,applied:next})});
    const d=await r.json();
    if(d.ok){ DATA.counts.applied=d.total; DATA.series=d.series; renderChart(d.series); renderStats(); }
  }catch(err){ /* keep optimistic state */ }
}
function dl(){window.location='/download?category='+CAT;}

async function scan(mode){
  document.getElementById('scanBtn').disabled=true;
  banner('<span class="spin"></span> Scanning your Telegram group… up to a minute.','');
  const r=await fetch('/api/scan?mode='+mode,{method:'POST'});
  if(r.status===409){banner('A scan is already running.','');return;}
  poll();
}
async function poll(){
  const s=await (await fetch('/api/status')).json();
  if(s.running){banner('<span class="spin"></span> '+esc(s.message||'Scanning…'),'');setTimeout(poll,1500);return;}
  document.getElementById('scanBtn').disabled=false;
  if(s.error){banner('❌ '+esc(s.error),'err');return;}
  banner('✅ '+(s.added!=null?s.added+' offers total. ':'')+'Updated.','');
  await load(); setTimeout(()=>banner(''),4000);
}
async function init(){
  await load();
  // If a scan is already running (auto catch-up on launch), show it and refresh when done.
  try{
    const s=await (await fetch('/api/status')).json();
    if(s.running){document.getElementById('scanBtn').disabled=true;
      banner('<span class="spin"></span> Auto-scanning for new offers…','');poll();}
  }catch(e){}
}
init();
</script>
</body></html>"""


def open_browser():
    webbrowser.open(f"http://127.0.0.1:{PORT}")


def _startup_scan():
    """On a cold start (e.g. first icon click after the PC was off), pull any
    messages that arrived while it was down, so the dashboard opens fresh."""
    if os.path.exists(pm.SESSION):
        _run_scan(full=False)


if __name__ == "__main__":
    if not os.path.exists(pm.SESSION):
        print("NOTE: not logged in yet. Run:  python placement_monitor.py --list")
    # The desktop-shortcut launcher opens the browser itself and passes --no-open.
    if "--no-open" not in sys.argv:
        threading.Timer(1.0, open_browser).start()
    # Kick off an automatic catch-up scan in the background as the server boots.
    threading.Thread(target=_startup_scan, daemon=True).start()
    print(f"\n  Placement Offers dashboard ->  http://127.0.0.1:{PORT}\n")
    app.run(port=PORT, debug=False)
