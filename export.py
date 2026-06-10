import json
import os
import glob
import datetime
import openpyxl

WORKBOOK = "WC 2026 Main model_vOPX.xlsx"
SHEET = "_Setup_PowerQuery"
SUBMISSIONS = "OPX submissions"
OUTPUT = "leaderboard.json"
GAMES_OUTPUT = "games.json"

# Group-stage scoring (matches the model): exact score = 3 (2 "exact" + 1 "1X2"),
# correct result only (W/D/L) = 1, wrong = 0. Correct group winner = 3 (handled
# separately, not per-game).
PTS_EXACT = 3
PTS_RESULT = 1


def sign(a, b):
    d = a - b
    return 0 if d == 0 else (1 if d > 0 else -1)


def game_points(phg, pag, hg, ag):
    if phg == hg and pag == ag:
        return PTS_EXACT
    if sign(phg, pag) == sign(hg, ag):
        return PTS_RESULT
    return 0


# ───────────────────────── leaderboard.json (summary) ─────────────────────────

# Previous ranks, so we can show movement since the last update.
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
    name = str(player).strip()
    pct = ws.cell(row=row_num, column=7).value
    rows.append({
        "rank":     ws.cell(row=row_num, column=1).value,
        "player":   name,
        "group":    ws.cell(row=row_num, column=3).value,
        "ko":       ws.cell(row=row_num, column=4).value,
        "bonus":    ws.cell(row=row_num, column=5).value,
        "total":    ws.cell(row=row_num, column=6).value,
        "pctMax":   round(float(pct), 4) if pct is not None else 0.0,
        "prevRank": prev_ranks.get(name),
    })
    row_num += 1

meta = {
    "lastRefresh": str(last_refresh) if last_refresh else "",
    "maxPoints":   max_points,
    "participants": participants,
}

with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump({"meta": meta, "rows": rows}, f, ensure_ascii=False, indent=2)

print(f"Exported {len(rows)} players -> {OUTPUT}")


# ───────────────────────── games.json (per-game detail) ─────────────────────────

# Actual group results + group winners from the Backend sheet.
# Columns (1-based): B=Group C=Home D=Away E=HomeGoals F=AwayGoals.
bk = wb["Backend"]
hdr = next(bk.iter_rows(min_row=2, max_row=2, values_only=True))
slot_i = hdr.index("Slot")
team_i = hdr.index("Team")

actual = {}            # (home, away) -> (hg, ag)
actual_winners = {}    # group letter -> winning team
for row in bk.iter_rows(min_row=3, values_only=True):
    grp, home, away, hg, ag = row[1], row[2], row[3], row[4], row[5]
    if grp and home and away and isinstance(hg, int) and isinstance(ag, int) and len(str(grp)) == 1:
        actual.setdefault((home, away), (hg, ag))
    slot = row[slot_i]
    if slot and str(slot).startswith("1") and len(str(slot)) == 2:
        actual_winners[str(slot)[1]] = row[team_i]

# Order players by leaderboard rank.
players = [r["player"] for r in sorted(rows, key=lambda r: (r["rank"] is None, r["rank"]))]

# Read predictions from each submission's VM-TIPS sheet.
# Columns (1-based): C=Home E=Away F=predHomeGoals H=predAwayGoals K=standingsPos L=standingsTeam.
schedule = {}          # (home, away) -> {"group": X, "kickoff": "..."}
preds = {}             # name -> {(home, away): (phg, pag)}
pred_winner = {}       # name -> {group: team}

for path in sorted(glob.glob(os.path.join(SUBMISSIONS, "*.xlsx"))):
    sub = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "VM-TIPS" not in sub.sheetnames:
        continue
    sh = sub["VM-TIPS"]
    name = sh["C2"].value
    if name is None:
        continue
    name = str(name).strip()
    preds.setdefault(name, {})
    pred_winner.setdefault(name, {})
    cur_group = None
    for row in sh.iter_rows(min_row=1, values_only=True):
        b = row[1]
        if isinstance(b, str) and b.startswith("GROUP"):
            cur_group = b.split()[-1]
        home, away, phg, pag = row[2], row[4], row[5], row[7]
        if home and away and isinstance(phg, int) and isinstance(pag, int):
            preds[name][(home, away)] = (phg, pag)
            if cur_group and (home, away) not in schedule:
                kickoff = b if isinstance(b, datetime.datetime) else None
                schedule[(home, away)] = {
                    "group": cur_group,
                    "kickoff": kickoff.strftime("%Y-%m-%d %H:%M") if kickoff else "",
                }
        # predicted group winner: K (col 11, index 10) == 1
        if cur_group and row[10] == 1 and row[11]:
            pred_winner[name][cur_group] = row[11]

# Build the per-group structure from the scheduled fixtures (the 72 real games).
groups_map = {}
for (home, away), info in schedule.items():
    groups_map.setdefault(info["group"], []).append((home, away, info["kickoff"]))

groups = []
for g in sorted(groups_map.keys()):
    fixtures = sorted(groups_map[g], key=lambda x: x[2])  # by kickoff
    games = []
    for home, away, kickoff in fixtures:
        res = actual.get((home, away))
        played = res is not None
        predictions = []
        for name in players:
            pred = preds.get(name, {}).get((home, away))
            entry = {"player": name}
            if pred is not None:
                entry["phg"], entry["pag"] = pred[0], pred[1]
                entry["pts"] = game_points(pred[0], pred[1], res[0], res[1]) if played else None
            else:
                entry["phg"] = entry["pag"] = entry["pts"] = None
            predictions.append(entry)
        games.append({
            "kickoff": kickoff,
            "home": home,
            "away": away,
            "played": played,
            "actual": {"hg": res[0], "ag": res[1]} if played else None,
            "predictions": predictions,
        })
    actual_w = actual_winners.get(g)
    winners = [{
        "player": name,
        "team": pred_winner.get(name, {}).get(g),
        "correct": (actual_w is not None and pred_winner.get(name, {}).get(g) == actual_w),
    } for name in players]
    groups.append({
        "group": g,
        "actualWinner": actual_w,
        "games": games,
        "winners": winners,
    })

with open(GAMES_OUTPUT, "w", encoding="utf-8") as f:
    json.dump({"meta": meta, "players": players, "groups": groups},
              f, ensure_ascii=False, indent=2)

ngames = sum(len(grp["games"]) for grp in groups)
print(f"Exported {ngames} games x {len(players)} players -> {GAMES_OUTPUT}")
