# FIFA World Cup 2026 — Live Leaderboard

A static leaderboard site for the OPX prediction pool.  
`export.py` reads the Excel model and writes `leaderboard.json`.  
`index.html` fetches that JSON and auto-refreshes every 60 seconds.

---

## Prerequisites

```
pip install openpyxl
```

---

## Re-run after each Excel update

1. Open a terminal in this folder.
2. Run:

```
python export.py
```

3. Verify `leaderboard.json` was created/updated (it prints the player count).
4. Commit and push (see below) — the site updates automatically within ~1 minute.

---

## Deploy to GitHub Pages

### First-time setup

1. Create a new **public** GitHub repository (e.g. `wc2026-leaderboard`).
2. In this folder, initialise git and push:

```bash
git init
git add export.py index.html leaderboard.json README.md
git commit -m "Initial leaderboard"
git branch -M main
git remote add origin https://github.com/<your-username>/wc2026-leaderboard.git
git push -u origin main
```

3. On GitHub → **Settings → Pages**:
   - Source: **Deploy from a branch**
   - Branch: `main` / `/ (root)`
   - Click **Save**

GitHub will give you a URL like `https://<your-username>.github.io/wc2026-leaderboard/`.

### After each Excel update

```bash
python export.py
git add leaderboard.json games.json
git commit -m "Update leaderboard $(date +%Y-%m-%d)"
git push
```

The live site reflects the new data within ~60 seconds.

> Tip: in Claude Code you can just run `/update-leaderboard-opx`, which does
> export → local preview → push for you (with a confirmation step before publishing).

---

## Local preview

Because `index.html` fetches `leaderboard.json` via `fetch()`, you need a local server (not `file://`):

```bash
python -m http.server 8000
# then open http://localhost:8000
```

---

## File overview

| File | Purpose |
|------|---------|
| `export.py` | Reads Excel + submissions, writes `leaderboard.json` and `games.json` |
| `leaderboard.json` | Summary standings, consumed by the Leaderboard tab (commit this) |
| `games.json` | Per-game points + everyone's predictions, for the Games tab (commit this) |
| `index.html` | Self-contained UI: Leaderboard + Games & Predictions tabs |
| `WC 2026 Main model_vOPX.xlsx` | Source of truth — do **not** commit |
| `OPX submissions/` | Players' prediction files — read by export.py, **not** committed |

### Games & Predictions tab

`export.py` reads each player's prediction file in `OPX submissions/`, compares
it to the actual results, and computes per-game points (exact score = 3, correct
result = 1, wrong = 0). The site shows all group games; tap one to see every
player's prediction and points. Predictions are always visible; the actual score
and points show once a game has been played.
