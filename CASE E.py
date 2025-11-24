import math
import pulp
from collections import defaultdict

# ------------ Config ----------
NURSE_CAPACITY_TYPE1 = 100
NURSE_CAPACITY_TYPE2 = 100
ROUND_UP = True   # ceil conversion from workload -> nurses
DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
SLOTS_3H = [
    "00:00-03:00","03:00-06:00","06:00-09:00","09:00-12:00",
    "12:00-15:00","15:00-18:00","18:00-21:00","21:00-24:00"
]

# Penalties (Table 9.18)
P_SHORTAGE = 100
P_PREFERENCE_WEIGHT = 1   # preference_score values will be multiplied by this
P_LEAVE = 100
P_SURPLUS = 20
CYCLICAL_PENALTY = 2
P_FAIRNESS = 5

STANDARD_SHIFTS_PER_WEEK = 6  # for fairness expectation

# ------------ Required workload per 3h slot (Dept A) - copy from your input ----------
required_workload_by_day = {
    'Monday'   : [320.7375817, 311.6095185, 407.0248565, 485.7720826, 420.8326040, 504.3048170, 444.9954619, 395.4970479],
    'Tuesday'  : [331.1048729, 321.8086302, 418.1746316, 498.6918502, 433.2898171, 517.5660400, 457.7510785, 408.1479951],
    'Wednesday': [351.5640668, 341.9868585, 443.6967895, 530.4620183, 460.8729339, 549.9066533, 486.5155754, 432.9819041],
    'Thursday' : [393.8539021, 382.6814528, 502.4045751, 601.1281482, 518.9735563, 623.8116058, 549.0822611, 485.5599442],
    'Friday'   : [374.6696302, 363.6192257, 481.8163405, 576.3303879, 495.2715800, 598.7660574, 525.0117155, 462.4014835],
    'Saturday' : [342.4186077, 332.7431552, 436.5435134, 522.3908920, 451.1360685, 542.0349924, 477.2321326, 422.1022322],
    'Sunday'   : [297.1598830, 288.5389601, 378.8946128, 452.2863628, 390.7353290, 469.7894487, 413.5996353, 366.6098370],
}

# ------------ Build min_coverage for 3h slots (ceil by default) ----------
def build_min_coverage_from_workload(capacity_per_nurse=100, roundup=True):
    min_cov = {}
    for d in DAYS:
        wk = required_workload_by_day[d]
        for i, slot in enumerate(SLOTS_3H):
            req = wk[i]
            if roundup:
                needed = math.ceil(req / capacity_per_nurse)
            else:
                needed = int(round(req / capacity_per_nurse))
            min_cov[(d, slot)] = needed
    return min_cov

# default min coverage using capacity=100 (both types equal)
min_coverage_3h = build_min_coverage_from_workload(100, ROUND_UP)

# ------------ Build nurse list from Exhibit 5 (basic parsing, simplified)
# We'll construct each nurse as a dict: id, type (1/2), employment (1.0/0.75), preferences text, fixed leaves set, cyc_total
# This is derived from the Exhibit 5 you pasted earlier.
nurse_rows = [
("301A001",1.0,1,"Preference for night shifts"),
("301A002",0.75,1,"Preference for early shifts"),
("301A003",1.0,1,"No early shifts"),
("301A004",1.0,1,"Preference for early shifts"),
("301A005",1.0,1,"Preference for late shifts"),
("301A006",1.0,1,""),
("301A007",1.0,1,""),
("301A008",1.0,1,"Preference for weekend day off"),
("301A009",1.0,1,""),
("301A010",1.0,1,"Preference for Wednesday day off"),
("302A011",0.75,2,""),
("302A012",0.75,2,"No weekend shifts"),
("302A013",1.0,2,"Preference for weekend day on"),
("302A014",0.75,2,""),
("302A015",0.75,2,""),
("302A016",0.75,2,"Preference for Tuesday day off, Thursday day on"),
("302A017",0.75,2,"Preference Monday day off"),
("302A018",0.75,2,""),
("302A019",0.75,2,"Preference for Thursday and Friday day off"),
("302A020",0.75,2,"Preference for weekend day off"),
("302A021",0.75,2,""),
("302A022",0.75,2,"Preference for weekend day on"),
("302A023",1.0,2,""),
("302A024",1.0,2,""),
("302A025",0.75,2,"Preference Monday day off"),
("302A026",0.75,2,""),
("302A027",0.75,2,"Preference for Thursday and Friday day off"),
("302A028",0.75,2,"Preference for weekend day off"),
("302A029",1.0,2,"No weekend shifts"),
("302A030",0.75,2,""),
("302A031",0.75,2,"Preference for weekend day on"),
("302A032",1.0,2,""),
]

nurses = []
for (nid, emp, ntype, preftext) in nurse_rows:
    fixed_leaves = set()
    # parse some fixed leave preferences heuristically:
    if "Wednesday day off" in preftext:
        fixed_leaves.add("Wednesday")
    if "Tuesday day off" in preftext:
        fixed_leaves.add("Tuesday")
    if "Monday day off" in preftext:
        fixed_leaves.add("Monday")
    if "Thursday and Friday day off" in preftext:
        fixed_leaves.update({"Thursday","Friday"})
    # No weekend shifts -> treat as fixed leave on weekend
    if "No weekend shifts" in preftext or "No weekend" in preftext:
        fixed_leaves.update({"Saturday","Sunday"})
    nurses.append({
        "id": nid, "employment": emp, "type": ntype,
        "preferences_text": preftext, "fixed_leaves": fixed_leaves
    })

# ------------ Build preference numeric score (1..9) for each (nurse,day,shift)
# Lower = better (preferred), 1 strong desired, 5 indifferent, 9 strong aversion, 100 = effectively forbidden
def build_preference_scores(nurses):
    pref = {}
    for n in nurses:
        txt = n['preferences_text'].lower()
        for d in DAYS:
            for s in SLOTS_3H:
                score = 5  # default indifferent
                # strong aversion if nurse declared "no early" and this 3h slot is a morning slot (06-12)
                if "no early" in txt and s in ["06:00-09:00","09:00-12:00"]:
                    score = 100
                if "preference for night" in txt or "preference for night shifts" in txt:
                    if s in ["21:00-24:00","00:00-03:00","03:00-06:00"]:  # treat night blocks as these
                        score = 1
                if "preference for early" in txt or "prefers early" in txt:
                    if s in ["06:00-09:00","09:00-12:00"]:
                        score = 1
                if "preference for late" in txt:
                    if s in ["12:00-15:00","15:00-18:00","18:00-21:00"]:
                        score = 1
                if "no weekend" in txt or "no weekend shifts" in txt:
                    if d in ["Saturday","Sunday"]:
                        score = 100
                # specific day-off preferences: heavy penalty if assigned those days
                if f"preference for {d} day off" in txt or f"preference {d} day off" in txt:
                    score = 100
                # weekend day on preference => prefer weekend blocks
                if "preference for weekend day on" in txt:
                    if d in ["Saturday","Sunday"]:
                        score = 1
                pref[(n['id'], d, s)] = score
    return pref

preference_scores = build_preference_scores(nurses)

# ------------ Utility to print roster nicely ----------
def print_roster(roster):
    # roster: list of tuples (nurse, day, slot)
    byday = {d: [] for d in DAYS}
    for (nid,d,s) in roster:
        byday[d].append((nid,s))
    for d in DAYS:
        print(f"\n{d}:")
        if not byday[d]:
            print("  <no assignments>")
            continue
        for (nid,s) in sorted(byday[d]):
            print(f"  {nid} -> {s}")

# ------------ Solver helper: build and run a PuLP model for 3h slots ----------
def solve_3h_roster(nurses, min_coverage, preference_scores,
                     force_cycle_soft=True, fairness=True, shortage_weight=P_SHORTAGE,
                     lexicographic=False):
    prob = pulp.LpProblem("3h_Roster", pulp.LpMinimize)
    # decision variables
    x = pulp.LpVariable.dicts("x", ((n['id'], d, s) for n in nurses for d in DAYS for s in SLOTS_3H), cat='Binary')
    short = pulp.LpVariable.dicts("short", ((d,s) for d in DAYS for s in SLOTS_3H), lowBound=0, cat='Integer')
    surp = pulp.LpVariable.dicts("surp", ((d,s) for d in DAYS for s in SLOTS_3H), lowBound=0, cat='Integer')
    cycdev = pulp.LpVariable.dicts("cycdev", (n['id'] for n in nurses), lowBound=0, cat='Integer')
    fairdev = pulp.LpVariable.dicts("fairdev", (n['id'] for n in nurses), lowBound=0, cat='Integer')

    # Objective: weighted sum
    obj = pulp.lpSum([shortage_weight * short[(d,s)] for d in DAYS for s in SLOTS_3H])
    obj += pulp.lpSum([P_SURPLUS * surp[(d,s)] for d in DAYS for s in SLOTS_3H])
    obj += pulp.lpSum([P_PREFERENCE_WEIGHT * preference_scores[(n['id'],d,s)] * x[(n['id'],d,s)]
                       for n in nurses for d in DAYS for s in SLOTS_3H])
    # leave violations as big penalties (we already encoded as preference 100 for those)
    if force_cycle_soft:
        obj += CYCLICAL_PENALTY * pulp.lpSum([cycdev[n['id']] for n in nurses])
    if fairness:
        obj += P_FAIRNESS * pulp.lpSum([fairdev[n['id']] for n in nurses])

    prob += obj

    # Constraints:
    # coverage / define short and surplus
    for d in DAYS:
        for s in SLOTS_3H:
            assigned = pulp.lpSum([x[(n['id'],d,s)] for n in nurses])
            prob += short[(d,s)] >= min_coverage.get((d,s), 0) - assigned
            prob += surp[(d,s)] >= assigned - min_coverage.get((d,s), 0)

    # one shift per nurse per day
    for n in nurses:
        for d in DAYS:
            prob += pulp.lpSum([x[(n['id'],d,s)] for s in SLOTS_3H]) <= 1

    # max shifts per nurse per week scaled by employment
    for n in nurses:
        max_shifts = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        prob += pulp.lpSum([x[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H]) <= max_shifts

    # night succession rules: if assigned night block, cannot be assigned early next day or late (we interpret early as morning block 06-12)
    for n in nurses:
        nid = n['id']
        for i in range(len(DAYS)-1):
            d = DAYS[i]; dnext = DAYS[i+1]
            # night blocks are 21-24 or 00-03; disallow next day's early blocks (06-12)
            prob += x[(nid,d,"21:00-24:00")] + x[(nid,dnext,"06:00-09:00")] <= 1
            prob += x[(nid,d,"21:00-24:00")] + x[(nid,dnext,"09:00-12:00")] <= 1

    # cyc deviation vs cyc baseline total slots S1+S2 (we only have cyc totals implicitly; approximate: use employment*standard)
    # For simplicity, we approximate baseline_total = round(employment*STANDARD_SHIFTS_PER_WEEK)
    for n in nurses:
        baseline_total = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        assigned_total = pulp.lpSum([x[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H])
        prob += cycdev[n['id']] >= assigned_total - baseline_total
        prob += cycdev[n['id']] >= baseline_total - assigned_total

    # fairness deviation
    for n in nurses:
        expected = n['employment'] * STANDARD_SHIFTS_PER_WEEK
        assigned_total = pulp.lpSum([x[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H])
        prob += fairdev[n['id']] >= assigned_total - expected
        prob += fairdev[n['id']] >= expected - assigned_total

    # enforce fixed leaves as hard constraints (if any)
    for n in nurses:
        for d in n['fixed_leaves']:
            for s in SLOTS_3H:
                prob += x[(n['id'],d,s)] == 0

    # Lexicographic option: first minimize shortages strictly
    if lexicographic:
        # Stage 1: minimize total shortages
        prob_stage1 = prob.copy()
        # Objective already includes shortage as primary if we set high weight; but we explicitly do two-stage:
        # Solve stage1 with objective of only shortages.
        only_short = pulp.lpSum([short[(d,s)] for d in DAYS for s in SLOTS_3H])
        prob_stage1.setObjective(only_short)
        prob_stage1.solve(pulp.PULP_CBC_CMD(msg=0))
        stage1_shortage = pulp.value(only_short)
        # fix total shortage to optimal value
        prob += pulp.lpSum([short[(d,s)] for d in DAYS for s in SLOTS_3H]) <= stage1_shortage
        # Now set the actual objective (preferences etc). Solve full problem
        prob.solve(pulp.PULP_CBC_CMD(msg=0))
    else:
        prob.solve(pulp.PULP_CBC_CMD(msg=0))

    # gather results
    roster = []
    for n in nurses:
        for d in DAYS:
            for s in SLOTS_3H:
                if pulp.value(x[(n['id'],d,s)]) >= 0.5:
                    roster.append((n['id'], d, s))
    shortages = {(d,s): int(pulp.value(short[(d,s)])) for d in DAYS for s in SLOTS_3H}
    surplus = {(d,s): int(pulp.value(surp[(d,s)])) for d in DAYS for s in SLOTS_3H}
    return roster, shortages, surplus

# ------------ Option 1: Basic 3h roster (capacity 100) ----------
print("\n=== OPTION 1: Basic 3-hour block roster (capacity 100) ===")
roster1, shortages1, surplus1 = solve_3h_roster(nurses, min_coverage_3h, preference_scores,
                                               force_cycle_soft=True, fairness=True, lexicographic=False)
print_roster(roster1)
print("\nNon-zero shortages (slot -> shortage):")
for k,v in shortages1.items():
    if v>0:
        print(k, "->", v)
print("\nNon-zero surplus (slot -> surplus):")
for k,v in surplus1.items():
    if v>0:
        print(k, "->", v)

# ------------ Option 2: Change capacity demonstration (kept as 100) ----------
print("\n=== OPTION 2: Change capacity (both types = 100) - same as option 1 (productivity constant) ===")
# Already both types =100 -> same min_coverage used. If you wanted to test capacity change, call build_min_coverage_from_workload with other value.
# Example (commented): min_coverage_alt = build_min_coverage_from_workload(capacity_per_nurse=120)
print("Both types productivity = 100 workload units per 3h block (confirmed).")

# ------------ Option 3: Aggregation into 9-hour shifts ----------
# We'll generate candidate 9h shifts starting at each 3h boundary: (start index 0..7) each 9h covers 3 consecutive 3h blocks.
def build_9h_candidates():
    candidates = []  # each candidate: (start_slot_index, list_of_3h_slots)
    for start_idx in range(0, 8):  # start at each 3h boundary
        slots = []
        for k in range(3):
            idx = (start_idx + k) % 8
            slots.append(SLOTS_3H[idx])
        candidates.append((start_idx, tuple(slots)))
    return candidates

cands_9h = build_9h_candidates()
# Build min coverage per day per candidate by summing required workload in those 3 blocks and dividing by capacity*3 (since one nurse in 9h provides 3 blocks * capacity)
def min_coverage_9h_from_3h(capacity_per_block=100, roundup=True):
    mincov = {}
    for d in DAYS:
        for (start_idx, blocks) in cands_9h:
            total_work = sum(required_workload_by_day[d][SLOTS_3H.index(b)] for b in blocks)
            # one 9h nurse provides 3 * capacity_per_block workload in that 9h
            cap_9h = 3 * capacity_per_block
            if roundup:
                needed = math.ceil(total_work / cap_9h)
            else:
                needed = int(round(total_work / cap_9h))
            mincov[(d, start_idx)] = needed
    return mincov

mincov_9h = min_coverage_9h_from_3h(100, ROUND_UP)

# Build a simple 9h LP: decision y[nurse,day,start_idx] binary (nurse works that 9h shift)
def solve_9h_roster(nurses, mincov_9h, preference_scores_3h):
    prob = pulp.LpProblem("9h_Roster", pulp.LpMinimize)
    # decision variables: y[(nid,day,start_idx)]
    y = pulp.LpVariable.dicts("y", ((n['id'],d,start_idx) for n in nurses for d in DAYS for start_idx in range(len(cands_9h))), cat='Binary')
    short = pulp.LpVariable.dicts("short", ((d,start_idx) for d in DAYS for start_idx in range(len(cands_9h))), lowBound=0, cat='Integer')
    surp  = pulp.LpVariable.dicts("surp", ((d,start_idx) for d in DAYS for start_idx in range(len(cands_9h))), lowBound=0, cat='Integer')

    # preference mapping: approximate 9h preference by summing 3h preference scores across the 3 blocks
    pref9 = {}
    for n in nurses:
        for d in DAYS:
            for idx, blocks in cands_9h:
                pref9[(n['id'],d,idx)] = sum(preference_scores[(n['id'],d,b)] for b in blocks)

    obj = pulp.lpSum([P_SHORTAGE * short[(d,idx)] for d in DAYS for idx in range(len(cands_9h))])
    obj += pulp.lpSum([P_SURPLUS * surp[(d,idx)] for d in DAYS for idx in range(len(cands_9h))])
    obj += pulp.lpSum([pref9[(n['id'],d,idx)] * y[(n['id'],d,idx)] for n in nurses for d in DAYS for idx,_ in cands_9h])
    prob += obj

    # coverage constraints
    for d in DAYS:
        for idx, blocks in cands_9h:
            assigned = pulp.lpSum([y[(n['id'],d,idx)] for n in nurses])
            prob += short[(d,idx)] >= mincov_9h[(d,idx)] - assigned
            prob += surp[(d,idx)] >= assigned - mincov_9h[(d,idx)]

    # each nurse can only do at most one 9h shift per day
    for n in nurses:
        for d in DAYS:
            prob += pulp.lpSum([y[(n['id'],d,idx)] for idx,_ in cands_9h]) <= 1
        # max shifts per week
        max_shifts = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        prob += pulp.lpSum([y[(n['id'],d,idx)] for d in DAYS for idx,_ in cands_9h]) <= max_shifts

    # fixed leaves
    for n in nurses:
        for d in n['fixed_leaves']:
            for idx,_ in cands_9h:
                prob += y[(n['id'],d,idx)] == 0

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    roster = []
    for n in nurses:
        for d in DAYS:
            for idx,_ in cands_9h:
                if pulp.value(y[(n['id'],d,idx)]) >= 0.5:
                    blocks = cands_9h[idx][1]
                    roster.append((n['id'], d, "9h_start_"+str(idx), blocks))
    shortages = {(d,idx): int(pulp.value(short[(d,idx)])) for d in DAYS for idx,_ in cands_9h}
    surplus = {(d,idx): int(pulp.value(surp[(d,idx)])) for d in DAYS for idx,_ in cands_9h}
    return roster, shortages, surplus

print("\n=== OPTION 3: 9-hour aggregated shifts ===")
roster9, short9, surp9 = solve_9h_roster(nurses, mincov_9h, preference_scores)
# print simple roster
for r in roster9[:30]:
    print(r)
print("\n9h shortages (non-zero):")
for k,v in short9.items():
    if v>0:
        print(k, "->", v)

# ------------ Option 4: Lexicographic two-stage solve for 3h roster ----------
print("\n=== OPTION 4: Lexicographic solve (minimize total shortages first, then preferences+fairness) ===")
# We'll implement two-stage explicitly:
# Stage 1: minimize total shortages
# Stage 2: fix total shortages to optimal and minimize remaining objective
def lexicographic_3h(nurses, min_coverage, preference_scores):
    # Stage1: build model with objective = sum shortages
    prob1 = pulp.LpProblem("Stage1_min_short", pulp.LpMinimize)
    x1 = pulp.LpVariable.dicts("x1", ((n['id'], d, s) for n in nurses for d in DAYS for s in SLOTS_3H), cat='Binary')
    short1 = pulp.LpVariable.dicts("short1", ((d,s) for d in DAYS for s in SLOTS_3H), lowBound=0, cat='Integer')
    # constraints
    for d in DAYS:
        for s in SLOTS_3H:
            assigned = pulp.lpSum([x1[(n['id'],d,s)] for n in nurses])
            prob1 += short1[(d,s)] >= min_coverage[(d,s)] - assigned
    for n in nurses:
        for d in DAYS:
            prob1 += pulp.lpSum([x1[(n['id'],d,s)] for s in SLOTS_3H]) <= 1
        max_shifts = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        prob1 += pulp.lpSum([x1[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H]) <= max_shifts
    # objective stage1
    prob1 += pulp.lpSum([short1[(d,s)] for d in DAYS for s in SLOTS_3H])
    prob1.solve(pulp.PULP_CBC_CMD(msg=0))
    opt_short = pulp.value(prob1.objective)
    print("Stage1 optimal total shortage:", opt_short)

    # Stage2: full objective with shortage fixed <= opt_short
    prob2 = pulp.LpProblem("Stage2_min_pref", pulp.LpMinimize)
    x2 = pulp.LpVariable.dicts("x2", ((n['id'], d, s) for n in nurses for d in DAYS for s in SLOTS_3H), cat='Binary')
    short2 = pulp.LpVariable.dicts("short2", ((d,s) for d in DAYS for s in SLOTS_3H), lowBound=0, cat='Integer')
    surp2 = pulp.LpVariable.dicts("surp2", ((d,s) for d in DAYS for s in SLOTS_3H), lowBound=0, cat='Integer')
    cycdev2 = pulp.LpVariable.dicts("cycdev2", (n['id'] for n in nurses), lowBound=0, cat='Integer')
    fairdev2 = pulp.LpVariable.dicts("fairdev2", (n['id'] for n in nurses), lowBound=0, cat='Integer')

    # constraints
    for d in DAYS:
        for s in SLOTS_3H:
            assigned = pulp.lpSum([x2[(n['id'],d,s)] for n in nurses])
            prob2 += short2[(d,s)] >= min_coverage[(d,s)] - assigned
            prob2 += surp2[(d,s)] >= assigned - min_coverage[(d,s)]
    for n in nurses:
        for d in DAYS:
            prob2 += pulp.lpSum([x2[(n['id'],d,s)] for s in SLOTS_3H]) <= 1
        max_shifts = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        prob2 += pulp.lpSum([x2[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H]) <= max_shifts
    for n in nurses:
        baseline_total = int(round(n['employment'] * STANDARD_SHIFTS_PER_WEEK))
        assigned_total = pulp.lpSum([x2[(n['id'],d,s)] for d in DAYS for s in SLOTS_3H])
        prob2 += cycdev2[n['id']] >= assigned_total - baseline_total
        prob2 += cycdev2[n['id']] >= baseline_total - assigned_total
        expected = n['employment'] * STANDARD_SHIFTS_PER_WEEK
        prob2 += fairdev2[n['id']] >= assigned_total - expected
        prob2 += fairdev2[n['id']] >= expected - assigned_total
        for d in n['fixed_leaves']:
            for s in SLOTS_3H:
                prob2 += x2[(n['id'],d,s)] == 0

    # fix shortage
    prob2 += pulp.lpSum([short2[(d,s)] for d in DAYS for s in SLOTS_3H]) <= opt_short

    # objective: preferences + surpluses + cycdev + fairness
    obj2 = pulp.lpSum([P_PREFERENCE_WEIGHT * preference_scores[(n['id'],d,s)] * x2[(n['id'],d,s)] for n in nurses for d in DAYS for s in SLOTS_3H])
    obj2 += pulp.lpSum([P_SURPLUS * surp2[(d,s)] for d in DAYS for s in SLOTS_3H])
    obj2 += CYCLICAL_PENALTY * pulp.lpSum([cycdev2[n['id']] for n in nurses])
    obj2 += P_FAIRNESS * pulp.lpSum([fairdev2[n['id']] for n in nurses])
    prob2 += obj2

    prob2.solve(pulp.PULP_CBC_CMD(msg=0))
    roster = []
    for n in nurses:
        for d in DAYS:
            for s in SLOTS_3H:
                if pulp.value(x2[(n['id'],d,s)]) >= 0.5:
                    roster.append((n['id'],d,s))
    shortages = {(d,s): int(pulp.value(short2[(d,s)])) for d in DAYS for s in SLOTS_3H}
    surplus = {(d,s): int(pulp.value(surp2[(d,s)])) for d in DAYS for s in SLOTS_3H}
    return roster, shortages, surplus

roster_lex, shorts_lex, surp_lex = lexicographic_3h(nurses, min_coverage_3h, preference_scores)
print_roster(roster_lex)
print("\nLex shortages (non-zero):")
for k,v in shorts_lex.items():
    if v>0:
        print(k, "->", v)
