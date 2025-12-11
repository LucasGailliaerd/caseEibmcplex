import math
import time
import random
import pandas as pd
from pathlib import Path
from copy import deepcopy

# ========== EXCEL PATH ==========
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "CASE_E_input.xlsx"

# ========== PROBLEM DIMENSIONS ==========
NURSES = 31
DAYS = 28
SHIFTS = 5
TYPES = 2

# 0 = F, 1 = E, 2 = D, 3 = L, 4 = N
SHIFT_LABELS = {0: "F", 1: "E", 2: "D", 3: "L", 4: "N"}
FREE_SHIFT = 0

# ========== WAGES ==========
# 0 = F (free), 1 = E, 2 = D, 3 = L, 4 = N
WAGE_WEEKDAY = {
    0: [0.0, 160.0, 160.0, 176.0, 216.0],
    1: [0.0, 120.0, 120.0, 132.0, 162.0],
}
WAGE_WEEKEND = {
    0: [0.0, 216.0, 216.0, 237.6, 291.6],
    1: [0.0, 162.0, 162.0, 178.2, 218.7],
}

# ========== OBJECTIVE WEIGHTS ==========
W_UNDER = 10000
W_OVER = 2000
W_ASSIGN = 2000
W_CONS = 1000
W_FORBID = 100000
W_MIN_ASSIGN = 8000  # or same order as W_ASSIGN


PREF_PEN = 20
LATE_EARLY_PEN = 1
NIGHT_REST_PEN = 1
CONTRACT_MIN_PEN = 0
CONS_WORK_PEN = 1
SHIFT_CHANGE_PEN = 5

WEIGHT_WAGE = 1
WEIGHT_NURSE = 1
WEIGHT_PATIENT = 1

# ========== GLOBAL PROBLEM STATE ==========
department = ""
number_days = 0
number_nurses = 0
number_shifts = 0
weekend = 7

hrs = [0] * SHIFTS
req = [[0] * SHIFTS for _ in range(DAYS)]
shift = [0] * SHIFTS
start_shift = [0] * SHIFTS
end_shift = [0] * SHIFTS

number_types = TYPES
nurse_type = [0] * NURSES
pref = [[[0] * SHIFTS for _ in range(DAYS)] for _ in range(NURSES)]
nurse_percent_employment = [0.0] * NURSES
personnel_number = [""] * NURSES

cyclic_roster = [[FREE_SHIFT] * DAYS for _ in range(NURSES)]
monthly_roster = [[FREE_SHIFT] * DAYS for _ in range(NURSES)]

min_ass = [0] * NURSES
max_ass = [0] * NURSES
identical = [0] * NURSES

max_cons = [[0] * SHIFTS for _ in range(NURSES)]
min_cons = [[0] * SHIFTS for _ in range(NURSES)]
min_shift = [[0] * SHIFTS for _ in range(NURSES)]
max_shift = [[0] * SHIFTS for _ in range(NURSES)]

min_cons_wrk = [0] * NURSES
max_cons_wrk = [0] * NURSES
extreme_max_cons = [[0] * SHIFTS for _ in range(NURSES)]
extreme_min_cons = [[0] * SHIFTS for _ in range(NURSES)]
extreme_max_cons_wrk = 0
extreme_min_cons_wrk = 0

count_shift = [0] * SHIFTS
scheduled = [[[0] * SHIFTS for _ in range(DAYS)] for _ in range(TYPES)]
violations = [0] * (DAYS * SHIFTS)

# ========== HELPERS ==========

def _find_cell_containing(df, text):
    target = text.upper()
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and target in val.strip().upper():
                return r, c
    raise ValueError(f"Cell containing '{text}' not found")


def _find_row_starting_with(df, text):
    target = text.upper()
    for r in range(df.shape[0]):
        val = df.iat[r, 0]
        if isinstance(val, str) and val.strip().upper().startswith(target):
            return r
    raise ValueError(f"Row starting with '{text}' not found")


def _find_col_with_label(df, row_idx, label):
    target = label.upper()
    for c in range(df.shape[1]):
        val = df.iat[row_idx, c]
        if isinstance(val, str) and val.strip().upper().startswith(target):
            return c
    raise ValueError(f"Column '{label}' not found in row {row_idx}")


def debug_list_sheets():
    xls = pd.ExcelFile(EXCEL_FILE)
    print("Excel file:", EXCEL_FILE)
    print("Sheets:", xls.sheet_names)


def is_weekend(day_idx: int) -> bool:
    d1 = day_idx + 1
    return (d1 % 7 == 6) or (d1 % 7 == 0)


def max_allowed_workblock(n: int) -> int:
    emp = nurse_percent_employment[n]
    if emp >= 0.99:
        return 5
    if emp >= 0.74:
        return 4
    return 3


def find_long_workblock(roster, n: int):
    if n >= len(roster):
        raise IndexError(f"Nurse index {n} out of bounds")
    max_len = max_allowed_workblock(n)
    cons = 0
    start = None
    for d in range(number_days + 1):
        if d < number_days and roster[n][d] != FREE_SHIFT:
            cons += 1
            if cons == 1:
                start = d
        else:
            if cons > max_len:
                return start, cons
            cons = 0
            start = None
    return None

def detect_number_nurses_from_monthly_roster():
    """
    Set global number_nurses based on non-empty 'Personnel Number'
    values in Case_E_MonthlyRoster_<department>.
    """
    global number_nurses

    sheet_name = f"Case_E_MonthlyRoster_{department}_EXTRA"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)

    if "Personnel Number" in df.columns:
        col = df["Personnel Number"].astype(str)
        mask = col.str.strip() != ""
        nurse_count = mask.sum()
    else:
        # Fallback: count non-empty rows in first column
        col = df.iloc[:, 0].astype(str)
        mask = col.str.strip() != ""
        nurse_count = mask.sum()

    if nurse_count == 0:
        raise ValueError(f"No personnel numbers found in sheet {sheet_name}.")

    if nurse_count > NURSES:
        raise ValueError(
            f"Monthly roster has {nurse_count} nurses but NURSES={NURSES} is the array capacity."
        )

    number_nurses = int(nurse_count)
    print(f"Detected number_nurses from {sheet_name}: {number_nurses}")

# ========== INPUT READING ==========

def read_shift_system():
    """
    Fill:
      - start_shift[], end_shift[]
      - hrs[shift_code]
      - req[day][shift_code]
    Using new encoding:
      1 = Early, 2 = Day, 3 = Late, 4 = Night, 0 = Free
    """
    global number_shifts, number_days
    sheet_name = "Case_C_9"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    number_shifts = int(df.iat[1, 0])  # number of working shifts (probably 4)
    length = int(df.iat[1, 1])

    r_start, c_start = _find_cell_containing(df, "START SHIFTS DEP B")
    r_req, c_req = _find_cell_containing(df, "REQUIREMENTS DEP B")

    start_rows = [r_start + 1 + i for i in range(number_shifts)]
    req_rows = [r_req + 1 + i for i in range(number_shifts)]

    for j in range(SHIFTS):
        req[0][j] = 0

    # working shift codes: 1..4
    for idx in range(number_shifts):
        row_start = start_rows[idx]
        row_req = req_rows[idx]

        start_h = int(df.iat[row_start, c_start])
        required = int(df.iat[row_req, c_req])

        k = idx + 1  # index in start_shift/end_shift arrays

        if 3 <= start_h < 9:      # Early
            code = 1
        elif 9 <= start_h < 12:   # Day
            code = 2
        elif 12 <= start_h < 21:  # Late
            code = 3
        else:                     # Night
            code = 4

        start_shift[k] = start_h
        shift[k] = code
        hrs[code] = length
        req[0][code] = required
        end_shift[k] = (start_h + length) % 24

    # free shift
    shift[0] = FREE_SHIFT
    hrs[FREE_SHIFT] = 0

    # copy day 0 requirements to all days
    for day in range(1, number_days):
        for j in range(SHIFTS):
            req[day][j] = req[0][j]

    number_shifts = SHIFTS

def read_personnel_characteristics():
    global number_nurses, number_types
    sheet_name = f"Case_E_Preferences_{department}_EXTRA"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)
    n_prefs = len(df)

    if n_prefs < number_nurses:
        raise ValueError(
            f"Preferences sheet {sheet_name} has only {n_prefs} rows, "
            f"but monthly roster indicates {number_nurses} nurses."
        )

    number_types = TYPES

    # 5 shifts per day, preferences in Excel are in order [E, D, L, N, F]
    prefs_per_nurse = SHIFTS * number_days  # 5 * number_days

    # Map "position inside the 5-column day block" -> internal shift code
    # Excel block: [E, D, L, N, F]
    # Internal:    0=F, 1=E, 2=D, 3=L, 4=N
    PREF_EXCEL_TO_INTERNAL = {
        0: 1,  # column B,G,... (Early)  -> 1
        1: 2,  # column C,H,... (Day)    -> 2
        2: 3,  # column D,I,... (Late)   -> 3
        3: 4,  # column E,J,... (Night)  -> 4
        4: 0,  # column F,K,... (Free)   -> 0
    }

    for k in range(number_nurses):
        row = df.iloc[k]

        # col A = personnel number
        personnel_number[k] = str(row.iloc[0])

        # cols B.. = preferences, flattened by row
        pref_values = row.iloc[1 : 1 + prefs_per_nurse].tolist()
        if len(pref_values) != prefs_per_nurse:
            raise ValueError(
                f"Row {k} in preferences sheet has {len(pref_values)} preference entries, "
                f"expected {prefs_per_nurse}."
            )

        # For each day, take a 5-column block [E,D,L,N,F]
        for day in range(number_days):
            start = day * SHIFTS
            day_block = pref_values[start : start + SHIFTS]  # length 5
            for excel_pos, val in enumerate(day_block):
                internal_shift = PREF_EXCEL_TO_INTERNAL[excel_pos]
                pref[k][day][internal_shift] = int(val)

        employment_col = 1 + prefs_per_nurse      # column after all prefs
        type_col = employment_col + 1

        nurse_percent_employment[k] = float(row.iloc[employment_col])
        nurse_type[k] = int(row.iloc[type_col]) - 1

def read_cyclic_roster():
    global number_nurses, number_days

    sheet_name = f"Case_D_Cyclic_{department}_EXTRA"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)

    n_cyc = len(df)

    if n_cyc < number_nurses:
        raise ValueError(
            f"Cyclic roster has only {n_cyc} rows, "
            f"but monthly roster indicates {number_nurses} nurses."
        )

    if "NurseType" not in df.columns:
        raise ValueError(f"'NurseType' column missing in sheet {sheet_name}")

    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(f"No Day* columns found in sheet {sheet_name}")

    excel_days = len(day_cols)
    if number_days != excel_days:
        print(
            f"WARNING: number_days in code = {number_days}, "
            f"but Excel has {excel_days} day columns. Using {excel_days}."
        )
        number_days = excel_days

    # Use only the first number_nurses rows (ignore extra ones)
    for k in range(number_nurses):
        nt_val = int(df.iloc[k]["NurseType"])
        nurse_type[k] = nt_val - 1
        for d_idx, col in enumerate(day_cols):
            excel_code = int(df.iloc[k][col])
            if excel_code < 0 or excel_code >= SHIFTS:
                raise ValueError(f"Unknown shift code {excel_code} in row {k+1}, column {col}")
            # Excel codes are identical to internal codes
            cyclic_roster[k][d_idx] = excel_code

def read_monthly_roster_constraints():
    global extreme_max_cons_wrk, extreme_min_cons_wrk

    sheet_name = "Case_E_Constraints_A"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    # total assignments
    r_ass = _find_row_starting_with(df, "NUMBER OF ASSIGNMENTS")
    header_row_ass = r_ass + 1
    val_row_ass = header_row_ass + 1

    min_col = _find_col_with_label(df, header_row_ass, "Minimum")
    max_col = _find_col_with_label(df, header_row_ass, "Maximum")

    base_min_ass = int(df.iat[val_row_ass, min_col])
    base_max_ass = int(df.iat[val_row_ass, max_col])

    # global consecutive assignments
    r_cons = _find_row_starting_with(df, "NUMBER OF CONSECUTIVE ASSIGNMENTS")
    for r in range(r_cons, df.shape[0]):
        val = df.iat[r, 0]
        if isinstance(val, str):
            txt = val.upper()
            if "NUMBER OF CONSECUTIVE ASSIGNMENTS" in txt and "PER SHIFT TYPE" not in txt:
                r_cons = r
                break

    header_row_cons = r_cons + 1
    val_row_cons = header_row_cons + 1

    min_col_cons = _find_col_with_label(df, header_row_cons, "Minimum")
    max_col_cons = _find_col_with_label(df, header_row_cons, "Maximum")

    base_min_cons_wrk = int(df.iat[val_row_cons, min_col_cons])
    base_max_cons_wrk = int(df.iat[val_row_cons, max_col_cons])

    # consecutive per shift type
    r_cons_sh = _find_row_starting_with(df, "NUMBER OF CONSECUTIVE ASSIGNMENTS PER SHIFT TYPE")
    header_row_cons_sh = r_cons_sh + 1
    first_val_row_cons_sh = header_row_cons_sh + 1

    min_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Minimum")
    max_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Maximum")

    base_min_cons = {}
    base_max_cons = {}

    working_shifts = [s for s in range(SHIFTS) if s != FREE_SHIFT]
    r = first_val_row_cons_sh
    for s_code in working_shifts:
        if r >= df.shape[0]:
            break
        val_min = df.iat[r, min_col_cons_sh]
        val_max = df.iat[r, max_col_cons_sh]
        if isinstance(val_min, (float, int)) and not pd.isna(val_min):
            base_min_cons[s_code] = int(val_min)
            base_max_cons[s_code] = int(val_max)
            r += 1
        else:
            break

    # assignments per shift type
    r_ass_sh = _find_row_starting_with(df, "NUMBER OF ASSIGNMENTS PER SHIFT TYPE")
    header_row_ass_sh = r_ass_sh + 1
    first_val_row_ass_sh = header_row_ass_sh + 1

    min_col_ass_sh = _find_col_with_label(df, header_row_ass_sh, "Minimum")
    max_col_ass_sh = _find_col_with_label(df, header_row_ass_sh, "Maximum")

    base_min_shift = {}
    base_max_shift = {}
    r = first_val_row_ass_sh
    for s_code in working_shifts:
        if r >= df.shape[0]:
            break
        val_min = df.iat[r, min_col_ass_sh]
        val_max = df.iat[r, max_col_ass_sh]
        if isinstance(val_min, (float, int)) and not pd.isna(val_min):
            base_min_shift[s_code] = int(val_min)
            base_max_shift[s_code] = int(val_max)
            r += 1
        else:
            break

    # identical weekend
    r_ident = _find_row_starting_with(df, "IDENTICAL WEEKEND CONSTRAINT")
    val_row_ident = r_ident + 1

    ident_value = None
    for c in range(df.shape[1]):
        cell = df.iat[val_row_ident, c]
        if isinstance(cell, str) and cell.strip():
            ident_value = cell.strip().upper()
            break
    ident_flag = 1 if (ident_value and ident_value.startswith("Y")) else 0

    for k in range(number_nurses):
        min_ass[k] = int(base_min_ass * nurse_percent_employment[k])
        max_ass[k] = int(base_max_ass * nurse_percent_employment[k])

        min_cons_wrk[k] = base_min_cons_wrk
        max_cons_wrk[k] = base_max_cons_wrk
        extreme_max_cons_wrk = 10
        extreme_min_cons_wrk = 1

        for s_code in working_shifts:
            min_cons[k][s_code] = base_min_cons.get(s_code, 0)
            max_cons[k][s_code] = base_max_cons.get(s_code, 28)
            extreme_max_cons[k][s_code] = 10
            extreme_min_cons[k][s_code] = 1
            min_shift[k][s_code] = base_min_shift.get(s_code, 0)
            max_shift[k][s_code] = base_max_shift.get(s_code, 9999)

        # free shift constraints
        min_cons[k][FREE_SHIFT] = 0
        max_cons[k][FREE_SHIFT] = 9999
        extreme_max_cons[k][FREE_SHIFT] = 10
        extreme_min_cons[k][FREE_SHIFT] = 1
        min_shift[k][FREE_SHIFT] = 0
        max_shift[k][FREE_SHIFT] = 9999

        identical[k] = ident_flag

def read_monthly_roster_from_excel():
    global number_days, number_nurses

    sheet_name = f"Case_E_MonthlyRoster_{department}_EXTRA"
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)

    # Recompute nurse count from personnel numbers
    if "Personnel Number" in df.columns:
        col = df["Personnel Number"].astype(str)
        mask = col.str.strip() != ""
        nurse_count = mask.sum()
    else:
        col = df.iloc[:, 0].astype(str)
        mask = col.str.strip() != ""
        nurse_count = mask.sum()

    if nurse_count != number_nurses:
        raise ValueError(
            f"Monthly roster nurse count ({nurse_count}) does not match "
            f"detected number_nurses ({number_nurses})."
        )

    # Optionally restrict df to these nurses (ignore extra blank rows)
    df = df[mask].reset_index(drop=True)

    # Consistency check with preferences
    if "Personnel Number" in df.columns:
        for k in range(number_nurses):
            roster_id = str(df.iloc[k]["Personnel Number"])
            if roster_id != personnel_number[k]:
                raise ValueError(
                    f"Mismatch between preferences and monthly roster at row {k}: "
                    f"prefs PN = {personnel_number[k]}, roster PN = {roster_id}"
                )

    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(f"No Day* columns found in sheet {sheet_name}")

    excel_days = len(day_cols)
    if excel_days != number_days:
        print(
            f"WARNING: code expects {number_days} days, "
            f"but Excel monthly roster has {excel_days} days. Using {excel_days}."
        )
        number_days = excel_days

    for k in range(number_nurses):
        for d_idx, col in enumerate(day_cols):
            excel_code = int(df.iloc[k][col])
            if excel_code < 0 or excel_code >= SHIFTS:
                raise ValueError(f"Unknown shift code {excel_code} for nurse {k+1}, day {d_idx+1}")
            monthly_roster[k][d_idx] = excel_code

# ========== DEBUG HELPERS ==========

def debug_print_first_nurse():
    print("\n=== FIRST NURSE CHECK ===")
    print(f"Personnel: {personnel_number[0]}")
    print(f"Employment %: {nurse_percent_employment[0]:.2f}")
    print(f"Type: {nurse_type[0] + 1}")
    print("Preferences (day 0):")
    for s in range(SHIFTS):
        print(f"  {SHIFT_LABELS.get(s, s)}: {pref[0][0][s]}")
    print(f"Assignments min={min_ass[0]}, max={max_ass[0]}")
    print(f"Consecutive working days: min={min_cons_wrk[0]}, max={max_cons_wrk[0]}")
    print("Consecutive shifts per type:")
    for s in range(SHIFTS):
        print(f"  {SHIFT_LABELS.get(s, s)}: min={min_cons[0][s]}, max={max_cons[0][s]}")
    print("Shift assignment limits:")
    for s in range(SHIFTS):
        print(f"  {SHIFT_LABELS.get(s, s)}: min={min_shift[0][s]}, max={max_shift[0][s]}")
    print(f"Identical weekend: {'YES' if identical[0] else 'NO'}")
    print("======================================\n")

def debug_capacity_vs_demand():
    working_shifts = [s for s in range(SHIFTS) if s != FREE_SHIFT]
    total_required = sum(
        req[d][s]
        for d in range(number_days)
        for s in working_shifts
    )
    total_max_assign = sum(max_ass[:number_nurses])
    total_min_assign = sum(min_ass[:number_nurses])

    print("Total required assignments:", total_required)
    print("Total min assignments:", total_min_assign)
    print("Total max assignments:", total_max_assign)

    if total_required > total_max_assign:
        print(">> INFEASIBLE: demand exceeds total max assignments.")
    elif total_required < total_min_assign:
        print(">> OVERCAPACITY: even min_ass exceeds demand.")
    else:
        print(">> Globally feasible in terms of total capacity.")

# ========== INPUT WRAPPER ==========

def read_input():
    global number_shifts
    read_shift_system()

    detect_number_nurses_from_monthly_roster()

    read_cyclic_roster()
    read_personnel_characteristics()
    read_monthly_roster_constraints()
    print("DEBUG nurse 1, day 1 prefs (F,E,D,L,N):")
    print([pref[0][0][s] for s in range(SHIFTS)])
    
    number_shifts = SHIFTS

# ========== OUTPUT WRITER ==========

def print_output():
    txt_filename = f"Monthly_Roster_dpt_{department}_EXTRA.txt"
    with open(txt_filename, "w") as f:
        for k in range(number_nurses):
            f.write(f"{personnel_number[k]}\t")
            for d in range(number_days):
                code = monthly_roster[k][d]
                f.write(f"{code}\t")
            f.write("\n")
    print(f"Monthly roster written to {txt_filename}")

    data = {"Personnel Number": [personnel_number[k] for k in range(number_nurses)]}
    for d in range(number_days):
        colname = f"Day{d + 1}"
        data[colname] = [monthly_roster[k][d] for k in range(number_nurses)]
    return pd.DataFrame(data)

# ========== EVALUATION ==========
def evaluate_line_of_work(nurse_idx: int, slack_j: int = 0):
    i = nurse_idx
    j = slack_j

    count_ass = 0
    count_cons_wrk = 0
    count_cons = 0
    for l in range(number_shifts):
        count_shift[l] = 0

    a = monthly_roster[i][0]
    violations[0] += pref[i][0][a]

    if a != FREE_SHIFT:
        count_ass += 1
        count_cons_wrk += 1
        count_cons += 1

    count_shift[a] += 1
    kk = nurse_type[i]
    scheduled[kk][0][a] += 1

    for k in range(1, number_days):
        h1 = monthly_roster[i][k]
        h2 = monthly_roster[i][k - 1]

        scheduled[kk][k][h1] += 1
        violations[0] += pref[i][k][h1]

        if h1 != FREE_SHIFT:
            count_ass += 1
            count_cons_wrk += 1
        count_shift[h1] += 1

        if h1 == FREE_SHIFT and h2 != FREE_SHIFT:
            if count_cons_wrk > max_cons_wrk[i] + j:
                violations[1] += 1
            count_cons_wrk = 0

        if h1 != h2:
            if h2 != FREE_SHIFT and count_cons > max_cons[i][h2] + j:
                violations[2] += 1
            count_cons = 1
        else:
            count_cons += 1

    if count_ass < min_ass[i]:
        violations[3] += 1
    if count_ass > max_ass[i]:
        violations[4] += 1


def evaluate_solution():
    for kk in range(number_types):
        for day in range(number_days):
            for sh in range(number_shifts):
                scheduled[kk][day][sh] = 0
    for idx in range(20):
        violations[idx] = 0

    for nurse_idx in range(number_nurses):
        evaluate_line_of_work(nurse_idx)

    txt_filename = BASE_DIR / f"Violations_dpt_{department}_EXTRA.txt"
    with open(txt_filename, "w") as f:
        f.write(f"The total preference score is {violations[0]}.\n")
        f.write(
            "The constraint 'maximum number of consecutive working days' "
            f"is violated {violations[1]} times.\n"
        )
        f.write(
            "The constraint 'maximum number of consecutive working days per shift type' "
            f"is violated {violations[2]} times.\n"
        )
        f.write(
            "The constraint 'minimum number of assignments' "
            f"is violated {violations[3]} times.\n"
        )
        f.write(
            "The constraint 'maximum number of assignments' "
            f"is violated {violations[4]} times.\n\n"
        )

        f.write("The staffing requirements are violated as follows:\n")
        working_shifts = [s for s in range(number_shifts) if s != FREE_SHIFT]
        for day in range(number_days):
            for sh in working_shifts:
                total_scheduled = sum(
                    scheduled[kk][day][sh] for kk in range(number_types)
                )
                required = req[day][sh]
                if total_scheduled < required:
                    f.write(
                        f"There are too few nurses in shift {sh} on day {day + 1}: "
                        f"{total_scheduled} < {required}.\n"
                    )
                elif total_scheduled > required:
                    f.write(
                        f"There are too many nurses in shift {sh} on day {day + 1}: "
                        f"{total_scheduled} > {required}.\n"
                    )

    print(f"Violations txt written to {txt_filename}")

    df_summary = pd.DataFrame([{
        "TotalPreferenceScore": violations[0],
        "MaxConsWorkViol": violations[1],
        "MaxConsShiftViol": violations[2],
        "MinAssignViol": violations[3],
        "MaxAssignViol": violations[4],
    }])

    staffing_rows = []
    working_shifts = [s for s in range(number_shifts) if s != FREE_SHIFT]
    for day in range(number_days):
        for sh in working_shifts:
            total_scheduled = sum(
                scheduled[kk][day][sh] for kk in range(number_types)
            )
            required = req[day][sh]
            if total_scheduled != required:
                staffing_rows.append({
                    "Day": day + 1,
                    "ShiftCode": sh,
                    "ShiftLabel": SHIFT_LABELS.get(sh, sh),
                    "Scheduled": total_scheduled,
                    "Required": required,
                    "Diff": total_scheduled - required,
                })
    df_staffing = pd.DataFrame(staffing_rows)
    return df_summary, df_staffing

# ========== COSTS / OBJECTIVE ==========

def count_consecutive_shifttype_violations(roster):
    total_viol = 0
    for i in range(number_nurses):
        count_cons = 0
        prev_s = roster[i][0]
        for k in range(1, number_days + 1):
            s = roster[i][k] if k < number_days else -1
            if s != prev_s:
                if prev_s != FREE_SHIFT and prev_s >= 0 and count_cons > max_cons[i][prev_s]:
                    total_viol += (count_cons - max_cons[i][prev_s])
                if s != FREE_SHIFT and s >= 0:
                    count_cons = 1
                else:
                    count_cons = 0
            else:
                if s != FREE_SHIFT and s >= 0:
                    count_cons += 1
            prev_s = s
    return total_viol


def compute_components(roster):
    wage_cost = 0.0
    nurse_cost = 0.0
    patient_cost = 0.0


    # wage + part of patient cost (shift changes)
    for n in range(number_nurses):
        works_anything = any(roster[n][d] != FREE_SHIFT for d in range(number_days))
        if not works_anything:
            continue

        t = nurse_type[n]
        for d in range(number_days):
            s = roster[n][d]
            if s != FREE_SHIFT:
                weekend_flag = is_weekend(d)
                table = WAGE_WEEKEND if weekend_flag else WAGE_WEEKDAY
                wage_cost += table[t][s]

    for n in range(number_nurses):
        works_anything = any(roster[n][d] != FREE_SHIFT for d in range(number_days))
        if not works_anything:
            continue

        for d in range(1, number_days):
            s_prev = roster[n][d - 1]
            s_curr = roster[n][d]
            if s_prev != FREE_SHIFT and s_curr != FREE_SHIFT and s_prev != s_curr:
                patient_cost += SHIFT_CHANGE_PEN

    # staffing under/over
    working_shifts = [s for s in range(number_shifts) if s != FREE_SHIFT]
    for d in range(number_days):
        for s in working_shifts:
            scheduled_count = sum(roster[n][d] == s for n in range(number_nurses))
            required = req[d][s]

            if required == 0 and scheduled_count > 0:
                patient_cost += W_FORBID * scheduled_count
                continue

            diff = scheduled_count - required
            if diff < 0:
                patient_cost += W_UNDER * ((-diff) ** 2)
            elif diff > 0:
                patient_cost += W_OVER * (diff ** 2)

    # nurse-side penalties
    EARLY_SHIFT = 1
    DAY_SHIFT = 2   # unused in rules but for clarity
    LATE_SHIFT = 3
    NIGHT_SHIFT = 4


    for n in range(number_nurses):
        works_anything = any(roster[n][d] != FREE_SHIFT for d in range(number_days))
        if not works_anything:
            continue

        # late-early / night-rest
        for d in range(1, number_days):
            prev_s = roster[n][d - 1]
            curr_s = roster[n][d]
            if prev_s == LATE_SHIFT and curr_s == EARLY_SHIFT:
                nurse_cost += LATE_EARLY_PEN
            if prev_s == NIGHT_SHIFT and curr_s in (EARLY_SHIFT, LATE_SHIFT):
                nurse_cost += NIGHT_REST_PEN

        # long working blocks
        limit = max_cons_wrk[n]
        if limit > 0:
            cons = 0
            for d in range(number_days + 1):
                if d < number_days and roster[n][d] != FREE_SHIFT:
                    cons += 1
                else:
                    if cons > limit:
                        nurse_cost += CONS_WORK_PEN * (cons - limit)
                    cons = 0

        # total assignments vs contracts
        worked = sum(roster[n][d] != FREE_SHIFT for d in range(number_days))
        emp = nurse_percent_employment[n]

        min_contract_shifts = 20 if emp >= 0.99 else 15

        # 1) Contractual minimum (hardish)
        if worked < min_contract_shifts:
            nurse_cost += W_MIN_ASSIGN * (min_contract_shifts - worked)

        # 2) Problem-specific min_ass (often same or similar)
        if worked < min_ass[n]:
            nurse_cost += W_MIN_ASSIGN * (min_ass[n] - worked)

        # 3) Over max assignments
        if worked > max_ass[n]:
            nurse_cost += W_ASSIGN * (worked - max_ass[n])

        # preferences
        for d in range(number_days):
            s = roster[n][d]
            if s != FREE_SHIFT:
                nurse_cost += PREF_PEN * pref[n][d][s]

    cons_shift_viol = count_consecutive_shifttype_violations(roster)
    nurse_cost += W_CONS * cons_shift_viol

    return wage_cost, nurse_cost, patient_cost


def violates_contract_min(roster) -> bool:
    for n in range(number_nurses):
        works_anything = any(roster[n][d] != FREE_SHIFT for d in range(number_days))
        if not works_anything:
            return True
        emp = nurse_percent_employment[n]
        worked = sum(roster[n][d] != FREE_SHIFT for d in range(number_days))
        min_contract_shifts = 20 if emp >= 0.99 else 15
        if worked < min_contract_shifts:
            return True
    return False

# ========== NEIGHBOR / SA ==========
def staffing_violation_score(roster):
    score = 0
    working_shifts = [s for s in range(number_shifts) if s != FREE_SHIFT]
    for d in range(number_days):
        for s in working_shifts:
            scheduled_count = sum(roster[n][d] == s for n in range(number_nurses))
            required = req[d][s]
            if required == 0 and scheduled_count > 0:
                score += scheduled_count  # forbidden shifts
            else:
                diff = scheduled_count - required
                score += abs(diff)
    return score

def random_neighbor(roster, p_swap=0.4, p_fix_block=0.3):
    new_roster = deepcopy(roster)

    if number_nurses < 1 or number_days < 1:
        return new_roster

    # try to break long work blocks first
    for n in random.sample(range(number_nurses), k=number_nurses):
        block = find_long_workblock(new_roster, n)
        if block is not None:
            start, length = block
            d = random.randint(start, start + length - 1)
            new_roster[n][d] = FREE_SHIFT
            return new_roster

    r = random.random()

    working_shifts = [s for s in range(SHIFTS) if s != FREE_SHIFT]

    if r < p_swap:
        # swap two nurses on a random day
        d = random.randrange(number_days)
        n1, n2 = random.sample(range(number_nurses), 2)
        new_roster[n1][d], new_roster[n2][d] = new_roster[n2][d], new_roster[n1][d]

    elif r < p_swap + p_fix_block:
        # try to fix consecutive-shift violations
        violation_found = False
        for n in random.sample(range(number_nurses), k=number_nurses):
            for s in working_shifts:
                limit = max_cons[n][s]
                if limit <= 0:
                    continue
                cons = 0
                start = None
                for d in range(number_days + 1):
                    if d < number_days and new_roster[n][d] == s:
                        cons += 1
                        if cons == 1:
                            start = d
                    else:
                        if cons > limit:
                            change_day = random.randint(start, start + cons - 1)
                            possible_shifts = [
                                x for x in range(SHIFTS)
                                if x != s and (req[change_day][x] > 0 or x == FREE_SHIFT)
                            ]
                            if possible_shifts:
                                understaffed = [
                                    x for x in possible_shifts
                                    if sum(new_roster[nn][change_day] == x for nn in range(number_nurses)) < req[change_day][x]
                                ]
                                new_roster[n][change_day] = random.choice(understaffed or possible_shifts)
                            violation_found = True
                            break
                        cons = 0
                if violation_found:
                    break
            if violation_found:
                break

        if not violation_found:
            n = random.randrange(number_nurses)
            d = random.randrange(number_days)
            old_shift = new_roster[n][d]
            possible_shifts = [
                x for x in range(SHIFTS)
                if x != old_shift and (req[d][x] > 0 or x == FREE_SHIFT)
            ]
            if possible_shifts:
                understaffed = [
                    x for x in possible_shifts
                    if sum(new_roster[nn][d] == x for nn in range(number_nurses)) < req[d][x]
                ]
                new_roster[n][d] = random.choice(understaffed or possible_shifts)

    else:
        n = random.randrange(number_nurses)
        d = random.randrange(number_days)
        old_shift = new_roster[n][d]
        possible_shifts = [
            x for x in range(SHIFTS)
            if x != old_shift and (req[d][x] > 0 or x == FREE_SHIFT)
        ]
        if possible_shifts:
            new_roster[n][d] = random.choice(possible_shifts)

    return new_roster


def simulated_annealing(initial_roster,
                        T_start=1000,
                        T_min=1,
                        alpha=0.95,
                        iters_per_T=200,
                        max_seconds = None):
    current = deepcopy(initial_roster)
    best = deepcopy(initial_roster)
    current_cost = compute_objective(current)
    best_cost = current_cost
    T = T_start

    start_time = time.perf_counter()

    while T > T_min:
        for _ in range(iters_per_T):
            if max_seconds is not None:
                elapsed = time.perf_counter() - start_time
                if elapsed >= max_seconds:
                    print(f"Time limit of {max_seconds} s reached, stopping SA.")
                    return best, best_cost
                
            neighbor = random_neighbor(current)
            neighbor_cost = compute_objective(neighbor)

            delta = neighbor_cost - current_cost

            if delta < 0:
                current, current_cost = neighbor, neighbor_cost
                if neighbor_cost < best_cost:
                    best, best_cost = deepcopy(neighbor), neighbor_cost
            else:
                p = math.exp(-delta / T)
                if random.random() < p:
                    current, current_cost = neighbor, neighbor_cost

        T *= alpha

    return best, best_cost

def add_nurse_to_day_shift(nurse_id: int, day_id: int, shift_id: int):
    monthly_roster[nurse_id][day_id] = shift_id


def compute_objective(roster):
    wage_cost, nurse_cost, patient_cost = compute_components(roster)
    return (
        WEIGHT_WAGE * wage_cost +
        WEIGHT_NURSE * nurse_cost +
        WEIGHT_PATIENT * patient_cost
    )

def procedure():
    global monthly_roster

    read_monthly_roster_from_excel()
    initial_roster = [
        [monthly_roster[n][d] for d in range(number_days)]
        for n in range(number_nurses)
    ]

    print("\n=== DEBUG: Staffing Day 1 BEFORE SA ===")
    print("Req per shift (day 1):")
    for s in range(SHIFTS):
        print(s, SHIFT_LABELS[s], "req =", req[0][s])

    print("\nScheduled per shift (day 1) BEFORE SA:")
    for s in range(SHIFTS):
        scheduled = sum(initial_roster[n][0] == s for n in range(number_nurses))
        print(s, SHIFT_LABELS[s], "scheduled =", scheduled)
    print("======================================\n")

    w0, n0, p0 = compute_components(initial_roster)
    obj0 = compute_objective(initial_roster)
    print("Initial schedule metrics:")
    print(f"  Wage_cost      = {w0:.2f}")
    print(f"  Nurse_cost     = {n0:.2f}")
    print(f"  Patient_cost   = {p0:.2f}")
    print(f"  Objective      = {obj0:.2f}")

    best_roster, best_obj = simulated_annealing(
        initial_roster,
        T_start=1000,
        T_min=1e-3,
        alpha=0.95,
        iters_per_T=200,
        max_seconds=300
    )

    w1, n1, p1 = compute_components(best_roster)
    print("Best schedule metrics after SA:")
    print(f"  Wage_cost      = {w1:.2f}")
    print(f"  Nurse_cost     = {n1:.2f}")
    print(f"  Patient_cost   = {p1:.2f}")
    print(f"  Objective      = {best_obj:.2f}")

    for n in range(number_nurses):
        for d in range(number_days):
            monthly_roster[n][d] = best_roster[n][d]

    print("First nurse, first 7 days (internal codes):")
    print(monthly_roster[0][:7])
    print("First nurse, first 7 days (labels):")
    print([SHIFT_LABELS[c] for c in monthly_roster[0][:7]])


def main():
    global number_days, weekend, department

    number_days = 28
    weekend = 7
    department = "B"

    seed = 42
    random.seed(seed)
    print(f"Using random seed: {seed}")

    debug_list_sheets()
    read_input()
    debug_capacity_vs_demand()
    debug_print_first_nurse()

    start_time = time.perf_counter()
    procedure()
    elapsed_time = time.perf_counter() - start_time
    print(f"CPU time for procedure(): {elapsed_time:.6f} seconds")

    df_roster = print_output()
    df_summary, df_staffing = evaluate_solution()

    output_file = BASE_DIR / f"CASE_E_output_{department}_EXTRA.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_roster.to_excel(writer, sheet_name="MonthlyRoster", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_staffing.to_excel(writer, sheet_name="StaffingViolations", index=False)

    print(f"\nExcel output written to: {output_file}")


if __name__ == "__main__":
    main()
