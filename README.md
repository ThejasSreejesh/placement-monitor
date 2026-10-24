# Placement Offer Monitor (Telegram → Excel)

Scans a Telegram group's **entire history**, detects placement / job offers,
extracts the key fields, shows them in a clean local dashboard, and exports a
tidy Excel workbook.

Uses **Telethon** (logs in as *your* Telegram account) because Telegram *bots*
cannot read past messages — only a user account can read full history.

## 🚀 Easiest way — the one-click dashboard

After the one-time login (see below), just **double-click
`Start Placement Dashboard.bat`**. Your browser opens a local dashboard where you:

- see stat tiles (total offers, ones with apply links / CTC, latest post date),
- browse offers, filter by date, click a card to expand the full message,
- **Scan new** / **Full rescan** with a button,
- **Download Excel** with a button.

No terminal needed after the first login. Close the black window to stop it.

---

## One-time login (needed once)

## 1. Install

```bash
cd placement_monitor
pip install -r requirements.txt
```

## 2. Get Telegram API credentials

1. Go to https://my.telegram.org → **API development tools**.
2. Create an app; copy the **api_id** and **api_hash**.

## 3. Configure

```bash
copy .env.example .env      # Windows
```

Edit `.env` and fill in `API_ID`, `API_HASH`, `PHONE`. Leave `GROUP` for now.

## 4. Find your group

```bash
python placement_monitor.py --list
```

First run sends a login code to your Telegram — enter it (and 2FA password if
you have one). Login is cached in `placement_monitor.session`, so later runs are
non-interactive. Copy the group's `@username`, `id`, or exact title into
`GROUP` in `.env`.

## 5. Run

```bash
python placement_monitor.py --once --full   # first time: scan whole history
python placement_monitor.py --once          # later: only new messages
python placement_monitor.py --schedule       # keep running, scan every N minutes
```

Output: **`placement_offers.xlsx`** — one row per offer, with a styled header,
frozen top row, autofilter, and a clickable Apply link. Duplicates are skipped
across runs via `state.json`.

## Columns

Date · Company · Role/Position · CTC/Stipend · Location · Eligibility/Batch ·
Apply Link · Posted By · Message

## Notes / tuning

- Detection is keyword-based (`OFFER_KEYWORDS` in the script) and field
  extraction reads `Label: value` lines plus a CTC regex. If your group posts in
  a particular format, tweak the `_guess_*` helpers — they're small and
  self-contained.
- To run truly unattended, use Windows **Task Scheduler** to call
  `python placement_monitor.py --once` on your interval instead of leaving
  `--schedule` running.
- Keep `.env` and the `.session` file private — they authenticate as you.
