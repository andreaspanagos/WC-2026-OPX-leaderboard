import json
import os
import openpyxl

WORKBOOK = "WC 2026 Main model_vOPX.xlsx"
SHEET = "_Setup_PowerQuery"
OUTPUT = "leaderboard.json"

# Read the previously published leaderboard (if any) so we can show how each
# player's rank changed since the last update. Maps player name -> old rank.
prev_ranks = {}
if os.path.exists(OUTPUT):
    try:
        with open(OUTPUT, encoding="utf-8") as f:
            old = json.load(f)
        for r in old.get("rows", []):
            if r.get("player") is not None:
                prev_ranks[r["player"]] = r.get("rank")
    except (json.JSONDecodeError, OSError):
        prev_ranks = {}

wb = openpyxl.load_workbook(WORKBOOK, data_only=True)
ws = wb[SHEET]

last_refresh = ws["B3"].value
max_points = ws["B4"].value
participants = ws["B5"].value

rows = []
row_num = 8
while True:
    player = ws.cell(row=row_num, column=2).value  # column B = player
    if player is None or str(player).strip() == "":
        break
    rank   = ws.cell(row=row_num, column=1).value
    group  = ws.cell(row=row_num, column=3).value
    ko     = ws.cell(row=row_num, column=4).value
    bonus  = ws.cell(row=row_num, column=5).value
    total  = ws.cell(row=row_num, column=6).value
    pct    = ws.cell(row=row_num, column=7).value
    name   = str(player).strip()
    rows.append({
        "rank":     rank,
        "player":   name,
        "group":    group,
        "ko":       ko,
        "bonus":    bonus,
        "total":    total,
        "pctMax":   round(float(pct), 4) if pct is not None else 0.0,
        "prevRank": prev_ranks.get(name),  # None => new entrant (shown as NEW)
    })
    row_num += 1

data = {
    "meta": {
        "lastRefresh": str(last_refresh) if last_refresh else "",
        "maxPoints":   max_points,
        "participants": participants,
    },
    "rows": rows,
}

with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"Exported {len(rows)} players -> {OUTPUT}")
