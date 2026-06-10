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
git add leaderboard.json games.json stats.json
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
| `export.py` | Reads the main model, writes `leaderboard.json`, `games.json`, `stats.json` |
| `leaderboard.json` | Summary standings — Leaderboard tab (commit this) |
| `games.json` | Per-game predicted-scoreline distributions + points — Games tab (commit this) |
| `stats.json` | Team progression predictions + bonus question status — Team Journeys & Bonus tabs (commit this) |
| `index.html` | Self-contained UI with all four tabs |
| `WC 2026 Main model_vOPX.xlsx` | Source of truth — do **not** commit |
| `OPX submissions/` | Players' original files — **not** committed, **not** read by export.py |

### How the site views work

`export.py` reads everything from the main model only (the flat per-player
picks table in `_Setup_PowerQuery` + results in `Backend`), so new players
appear automatically once they're added to the model — no extra steps.

- **Games** — every group game with the pool's predicted scorelines grouped
  and sorted by popularity, plus points per scoreline once the game is played
  (exact = 3, correct outcome = 1). Group headers show winner-pick percentages.
- **Team Journeys** — for each nation, the share of the pool predicting it to
  exit at each stage (group / R32 / R16 / QF / SF / Final / Champion), plus how
  far the team has actually come.
- **Bonus Questions** — everyone's answers per question and the currently
  correct answer where it can be derived from results ("so far" = may change,
  "TBD" = needs manual facts like top scorer or cards). Official bonus points
  remain whatever the model says on the leaderboard.

At tournament start the Results sheet is empty: all games show as *upcoming*,
no winners or points are shown, and bonus answers are TBD. As results are
entered, everything fills in automatically on the next export.
