"""Export website data from the WC 2026 main model.

Reads ONLY the main model workbook (not the individual submission files):
  - `_Setup_PowerQuery` A7-table  -> leaderboard.json   (official standings)
  - `_Setup_PowerQuery` row-100 flat table (one row per player, all picks)
  - `Backend` sheet               -> fixtures, results, per-team progression
  - `Scoring` sheet               -> bonus question texts + point values

Outputs:
  leaderboard.json  summary standings (+ rank movement vs previous publish)
  games.json        per-game predicted-scoreline distributions + points
  stats.json        per-team predicted progression + bonus question status

Scoring (validated against the model's own totals):
  group game: exact score = 3, correct 1X2 = 1, wrong = 0
  group winner = 3; KO "team reaches round": R32=3 R16=6 QF=9 SF=12 Final=15,
  champion = 30, 3rd place = 15.
All readers are defensive about missing results: at tournament start the
Results sheet is empty -> games show as upcoming, no winners, bonus mostly TBD.
"""

import json
import os
import openpyxl

WORKBOOK = "WC 2026 Main model_vOPX.xlsx"
SHEET = "_Setup_PowerQuery"
OUTPUT = "leaderboard.json"
GAMES_OUTPUT = "games.json"
STATS_OUTPUT = "stats.json"

FLAT_HEADER_ROW = 100          # header row of the per-player flat picks table
HOSTS = {"USA", "Mexico", "Canada"}
DEBUTANTS = {"Curaçao", "Cabo Verde", "Uzbekistan", "Jordan"}
STAGES = ["Group stage", "R32", "R16", "QF", "SF", "Final", "Champion"]


def sign(a, b):
    d = a - b
    return 0 if d == 0 else (1 if d > 0 else -1)


def game_points(phg, pag, hg, ag):
    if phg == hg and pag == ag:
        return 3
    if sign(phg, pag) == sign(hg, ag):
        return 1
    return 0


def as_int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().lstrip("-").isdigit():
        return int(v.strip())
    return None


wb = openpyxl.load_workbook(WORKBOOK, data_only=True)
ws = wb[SHEET]
bk = wb["Backend"]

# ───────────────────────── leaderboard.json ─────────────────────────

prev_ranks = {}
if os.path.exists(OUTPUT):
    try:
        with open(OUTPUT, encoding="utf-8") as f:
            for r in json.load(f).get("rows", []):
                if r.get("player") is not None:
                    prev_ranks[r["player"]] = r.get("rank")
    except (json.JSONDecodeError, OSError):
        prev_ranks = {}

meta = {
    "lastRefresh": str(ws["B3"].value) if ws["B3"].value else "",
    "maxPoints": ws["B4"].value,
    "participants": ws["B5"].value,
}

rows = []
row_num = 8
while True:
    player = ws.cell(row=row_num, column=2).value
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

with open(OUTPUT, "w", encoding="utf-8") as f:
    json.dump({"meta": meta, "rows": rows}, f, ensure_ascii=False, indent=2)
print(f"Exported {len(rows)} players -> {OUTPUT}")

# ───────────────── flat per-player picks table (row 100+) ─────────────────

hdr = {}
for c in range(1, ws.max_column + 1):
    h = ws.cell(row=FLAT_HEADER_ROW, column=c).value
    if h is not None:
        hdr[str(h)] = c

flat = {}                      # player name -> {header: value}
r = FLAT_HEADER_ROW + 1
while True:
    name = ws.cell(row=r, column=hdr["Name"]).value
    if name is None or str(name).strip() == "":
        break
    flat[str(name).strip()] = {h: ws.cell(row=r, column=c).value for h, c in hdr.items()}
    r += 1

# Players ordered by leaderboard rank; fall back to flat-table order.
players = [x["player"] for x in sorted(rows, key=lambda x: (x["rank"] is None, x["rank"]))
           if x["player"] in flat]
players += [n for n in flat if n not in players]

# ───────────────── Backend: fixtures, results, progression ─────────────────

# Group fixtures in GS01..GS72 order (Backend rows 3-74, cols B-F).
fixtures = []                  # [{group, home, away, hg, ag, played}]
for row in bk.iter_rows(min_row=3, max_row=74, min_col=2, max_col=6, values_only=True):
    grp, home, away, hg, ag = row
    if grp and home and away and len(str(grp)) == 1:
        hg, ag = as_int(hg), as_int(ag)
        fixtures.append({"group": str(grp), "home": home, "away": away,
                         "hg": hg, "ag": ag, "played": hg is not None and ag is not None})

# Kickoff datetimes from the Template sheet (same fixture order per group).
kickoffs = {}
tpl = wb["Template"] if "Template" in wb.sheetnames else None
if tpl is not None:
    import datetime as _dt
    for row in tpl.iter_rows(min_row=1, max_col=8, values_only=True):
        b, home, away = row[1], row[2], row[4]
        if isinstance(b, _dt.datetime) and home and away:
            kickoffs.setdefault((home, away), b.strftime("%Y-%m-%d %H:%M"))

# Per-group completeness flags (Backend AO/AP, rows 2-13).
group_complete = {}
for row in bk.iter_rows(min_row=2, max_row=13, min_col=41, max_col=42, values_only=True):
    g, done = row
    if g and len(str(g)) == 1:
        group_complete[str(g)] = str(done) == "True"

# Team table (Backend cols P-Z from row 3): slot, team, group + outcome flags.
teams = {}                     # team -> {"group": g, "flags": {...}}
actual_winner = {}             # group -> winning team (only if group complete)
for row in bk.iter_rows(min_row=3, max_row=50, min_col=16, max_col=26, values_only=True):
    slot, team, grp = row[0], row[1], row[2]
    if not team:
        continue
    flags = dict(zip(["R32", "R16", "QF", "SF", "Final", "Third", "RunnerUp", "Winner"],
                     [str(v) == "True" for v in row[3:11]]))
    teams[team] = {"group": str(grp) if grp else None, "flags": flags}
    if slot and str(slot).startswith("1") and len(str(slot)) == 2 and group_complete.get(str(slot)[1]):
        actual_winner[str(slot)[1]] = team

# KO rounds by match participation (Backend cols I-N from row 3).
# NOTE: the team table's "Quarter" flag column is broken in the model, so
# round membership is derived from the KO match list instead (validated).
ko_round = {"R32": set(), "R16": set(), "QF": set(), "SF": set(), "Final": set()}
round_map = {"R32": "R32", "R16": "R16", "QF": "QF", "SF": "SF", "Final": "Final"}
for row in bk.iter_rows(min_row=3, max_row=40, min_col=9, max_col=14, values_only=True):
    m, rnd, home, away, hg, ag = row
    key = round_map.get(str(rnd)) if rnd else None
    if key:
        for t in (home, away):
            if t in teams:
                ko_round[key].add(t)

champion = {t for t, d in teams.items() if d["flags"]["Winner"]}
third_team = {t for t, d in teams.items() if d["flags"]["Third"]}

def actual_stage(team):
    """Deepest stage the team has reached so far, or None before the R32 draw."""
    if team in champion:
        return "Champion"
    for key in ["Final", "SF", "QF", "R16", "R32"]:
        if team in ko_round[key]:
            return key
    if ko_round["R32"] and group_complete.get(teams[team]["group"]):
        return "Group stage"   # group done, R32 drawn, team not in it -> eliminated
    return None

# ───────────────── per-player derived picks ─────────────────

def player_round_sets(p):
    d = flat[p]
    rq  = {d.get(f"RQ_{i:02d}") for i in range(1, 33)}
    r16 = {d.get(f"R16_{i:02d}") for i in range(1, 17)}
    qf  = {d.get(f"QF_{i:02d}") for i in range(1, 9)}
    sf  = {d.get(f"SF_{i:02d}") for i in range(1, 5)}
    fin = {d.get("FIN_1"), d.get("FIN_2")}
    return {"R32": rq, "R16": r16, "QF": qf, "SF": sf, "Final": fin,
            "Winner": d.get("Winner"), "Third": d.get("Third")}

rounds_by_player = {p: player_round_sets(p) for p in players}

def predicted_stage(p, team):
    rs = rounds_by_player[p]
    if rs["Winner"] == team:
        return "Champion"
    for key in ["Final", "SF", "QF", "R16", "R32"]:
        if team in rs[key]:
            return key
    return "Group stage"

# ───────────────── games.json: scoreline distributions ─────────────────

n_players = len(players)
groups_out = {}
for gi, fx in enumerate(fixtures):
    gs = f"GS{gi + 1:02d}"
    by_score = {}
    no_pick = []
    for p in players:
        ph, pa = as_int(flat[p].get(gs + "_H")), as_int(flat[p].get(gs + "_A"))
        if ph is None or pa is None:
            no_pick.append(p)
            continue
        by_score.setdefault((ph, pa), []).append(p)
    scorelines = []
    for (ph, pa), ppl in by_score.items():
        scorelines.append({
            "score": f"{ph}–{pa}",
            "count": len(ppl),
            "pct": round(100 * len(ppl) / n_players) if n_players else 0,
            "pts": game_points(ph, pa, fx["hg"], fx["ag"]) if fx["played"] else None,
            "players": ppl,
        })
    scorelines.sort(key=lambda s: (-s["count"], s["score"]))
    groups_out.setdefault(fx["group"], []).append({
        "kickoff": kickoffs.get((fx["home"], fx["away"]), ""),
        "home": fx["home"], "away": fx["away"],
        "played": fx["played"],
        "actual": {"hg": fx["hg"], "ag": fx["ag"]} if fx["played"] else None,
        "scorelines": scorelines,
        "noPick": len(no_pick),
    })

groups_json = []
for g in sorted(groups_out):
    winner_counts = {}
    for p in players:
        t = flat[p].get(f"GW_{g}")
        if t:
            winner_counts.setdefault(t, []).append(p)
    picks = [{
        "team": t,
        "count": len(ppl),
        "pct": round(100 * len(ppl) / n_players) if n_players else 0,
        "players": ppl,
        "correct": actual_winner.get(g) == t if g in actual_winner else None,
    } for t, ppl in winner_counts.items()]
    picks.sort(key=lambda x: (-x["count"], x["team"]))
    groups_json.append({
        "group": g,
        "complete": group_complete.get(g, False),
        "actualWinner": actual_winner.get(g),
        "winnerPicks": picks,
        "games": groups_out[g],
    })

with open(GAMES_OUTPUT, "w", encoding="utf-8") as f:
    json.dump({"meta": meta, "players": players, "groups": groups_json},
              f, ensure_ascii=False, indent=2)
ngames = sum(len(g["games"]) for g in groups_json)
print(f"Exported {ngames} games x {n_players} players -> {GAMES_OUTPUT}")

# ───────────────── stats.json: team progression + bonus ─────────────────

teams_json = []
for team in sorted(teams, key=lambda t: (teams[t]["group"] or "?", t)):
    dist = {s: 0 for s in STAGES}
    champions_by = []
    for p in players:
        st = predicted_stage(p, team)
        dist[st] += 1
        if st == "Champion":
            champions_by.append(p)
    teams_json.append({
        "team": team,
        "group": teams[team]["group"],
        "actual": actual_stage(team),
        "dist": dist,
        "pct": {s: (round(100 * c / n_players) if n_players else 0) for s, c in dist.items()},
        "championPct": round(100 * dist["Champion"] / n_players) if n_players else 0,
        "championBy": champions_by,
    })

# Bonus questions: texts/points from Scoring rows 19-33.
sc = wb["Scoring"]
bonus_defs = []
for row in sc.iter_rows(min_row=19, max_row=33, min_col=2, max_col=4, values_only=True):
    bid, q, pts = row
    if bid and q:
        bonus_defs.append({"id": str(bid), "q": str(q), "pts": pts})

# Current correct answers, derived from results where possible.
all_groups_done = bool(group_complete) and all(group_complete.get(g, False)
                                               for g in group_complete)
played_group = [f for f in fixtures if f["played"]]
ko_played = []
for row in bk.iter_rows(min_row=3, max_row=40, min_col=9, max_col=14, values_only=True):
    m, rnd, home, away, hg, ag = row
    if rnd and as_int(hg) is not None and as_int(ag) is not None:
        ko_played.append((str(rnd), home, away, as_int(hg), as_int(ag)))

# Goals scored/conceded in group stage per team (from played group games).
gs_for, gs_against = {}, {}
for fxx in played_group:
    gs_for[fxx["home"]] = gs_for.get(fxx["home"], 0) + fxx["hg"]
    gs_for[fxx["away"]] = gs_for.get(fxx["away"], 0) + fxx["ag"]
    gs_against[fxx["home"]] = gs_against.get(fxx["home"], 0) + fxx["ag"]
    gs_against[fxx["away"]] = gs_against.get(fxx["away"], 0) + fxx["hg"]

total_goals = sum(f["hg"] + f["ag"] for f in played_group) + \
              sum(hg + ag for _, _, _, hg, ag in ko_played)
all_scores = [f["hg"] + f["ag"] for f in played_group] + \
             [hg + ag for _, _, _, hg, ag in ko_played]
sweden_now = actual_stage("Sweden") if "Sweden" in teams else None

def maxima(d):
    if not d:
        return []
    m = max(d.values())
    return sorted(t for t, v in d.items() if v == m)

def yes_no_reached(group_of_teams, round_set, round_size):
    hit = bool(group_of_teams & round_set)
    if hit:
        return "Yes", "decided"
    if len(round_set) >= round_size:
        return "No", "decided"
    return None, "tbd"

b11_ans, b11_st = yes_no_reached(HOSTS, ko_round["R16"], 16)
b12_ans, b12_st = yes_no_reached(DEBUTANTS, ko_round["R32"], 32)

tournament_over = bool(champion)
all_played = len(played_group) == 72 and len(ko_played) == 32

# Sweden's run is settled once she's eliminated (next round fully drawn
# without her) or the tournament is over.
ROUND_SIZE = {"R32": 32, "R16": 16, "QF": 8, "SF": 4, "Final": 2}
def sweden_settled():
    if "Sweden" not in teams or sweden_now is None:
        return False
    if tournament_over or sweden_now == "Group stage":
        return True
    nxt = {"R32": "R16", "R16": "QF", "QF": "SF", "SF": "Final"}.get(sweden_now)
    return (nxt is not None and len(ko_round[nxt]) >= ROUND_SIZE[nxt]
            and "Sweden" not in ko_round[nxt])

# `current` values may be a list (ties allowed); hit = membership.
current = {
    "B01": (maxima(gs_for) or None,
            "decided" if all_groups_done else ("provisional" if gs_for else "tbd")),
    "B02": (maxima(gs_against) or None,
            "decided" if all_groups_done else ("provisional" if gs_against else "tbd")),
    "B03": (total_goals if all_scores else None,
            "decided" if all_played else ("provisional" if all_scores else "tbd")),
    "B08": (max(all_scores) if all_scores else None,
            "decided" if all_played else ("provisional" if all_scores else "tbd")),
    "B11": (b11_ans, b11_st),
    "B12": (b12_ans, b12_st),
    "B15": (sweden_now, "decided" if sweden_settled() else ("provisional" if sweden_now else "tbd")),
}

# Per-player answers per question.
def bonus_answer(p, bid):
    if bid == "B15":
        return predicted_stage(p, "Sweden")
    if bid == "B16":
        return flat[p].get("B16_TopScorer")
    return flat[p].get(bid)

bonus_json = []
for bd in bonus_defs:
    bid = bd["id"]
    cur, status = current.get(bid, (None, "tbd"))
    cur_set = {str(c) for c in cur} if isinstance(cur, list) else \
              ({str(cur)} if cur is not None else set())
    counts = {}
    for p in players:
        a = bonus_answer(p, bid)
        a = "—" if a is None or str(a).strip() == "" else str(a).strip()
        counts.setdefault(a, []).append(p)
    answers = [{
        "answer": a,
        "count": len(ppl),
        "pct": round(100 * len(ppl) / n_players) if n_players else 0,
        "players": ppl,
        "hit": (a in cur_set) if (status == "decided" and cur_set) else None,
    } for a, ppl in counts.items()]
    answers.sort(key=lambda x: (-x["count"], x["answer"]))
    cur_out = ", ".join(str(c) for c in cur) if isinstance(cur, list) else cur
    bonus_json.append({**bd, "current": cur_out, "status": status, "answers": answers})

with open(STATS_OUTPUT, "w", encoding="utf-8") as f:
    json.dump({"meta": meta, "players": players, "stages": STAGES,
               "teams": teams_json, "bonus": bonus_json},
              f, ensure_ascii=False, indent=2)
print(f"Exported {len(teams_json)} teams, {len(bonus_json)} bonus questions -> {STATS_OUTPUT}")

# ───────────────── self-check vs official leaderboard ─────────────────

ko_pts_map = {"R32": 3, "R16": 6, "QF": 9, "SF": 12, "Final": 15}
for lbrow in rows:
    p = lbrow["player"]
    if p not in flat:
        continue
    gpts = 0
    for gi, fx in enumerate(fixtures):
        if not fx["played"]:
            continue
        ph = as_int(flat[p].get(f"GS{gi + 1:02d}_H"))
        pa = as_int(flat[p].get(f"GS{gi + 1:02d}_A"))
        if ph is not None and pa is not None:
            gpts += game_points(ph, pa, fx["hg"], fx["ag"])
    for g, w in actual_winner.items():
        if flat[p].get(f"GW_{g}") == w:
            gpts += 3
    rs = rounds_by_player[p]
    kpts = sum(len(rs[k] & ko_round[k]) * v for k, v in ko_pts_map.items())
    kpts += 30 if rs["Winner"] in champion else 0
    kpts += 15 if rs["Third"] in third_team else 0
    if lbrow["group"] is not None and gpts != lbrow["group"]:
        print(f"WARNING: {p} computed group pts {gpts} != leaderboard {lbrow['group']}")
    if lbrow["ko"] is not None and kpts != lbrow["ko"]:
        print(f"WARNING: {p} computed KO pts {kpts} != leaderboard {lbrow['ko']}")
