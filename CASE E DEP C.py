import math
import time
import random
import pandas as pd
from pathlib import Path
from copy import deepcopy

# Paths and constants

BASE_DIR = Path(__file__).resolve().parent
excel_file = BASE_DIR / "CASE_E_input.xlsx"

SHIFT_LABELS = {0: "E", 1: "D", 2: "L", 3: "N", 4: "F"}
FREE_SHIFT = max(SHIFT_LABELS.keys())

def shift_decoding(code: int) -> str:
    """
    Decode a shift code into a descriptive string:
    Example: 'E (start 06, end 14, hrs 8)'.
    """
    label = SHIFT_LABELS.get(code, "?")

    if code == 4:
        return "F (Free)"

    # shifts 1–3 map to indexes 1–3 in start_shift[], end_shift[]
    # internal code 0..3 maps to k = 1..4 but only first N shifts matter
    # Use inverse lookup:
    k = None
    for idx in range(1, number_shifts + 1):
        if shift[idx] == code:
            k = idx
            break

    if k is None:
        return f"{label} (unknown shift definition)"

    start = start_shift[k]
    end = end_shift[k]
    hours = hrs[code]

    return f"{label} (start {start:02d}, end {end:02d}, hrs {hours})"


# Objective weights
W_PREF   = 20    # nurse dissatisfaction (preference score)
W_UNDER  = 5000   # penalty per nurse missing (understaffing)
W_OVER   = 500   # penalty per nurse extra (overstaffing)
W_ASSIGN = 10000    # penalty per shifts beyond min/max total assignments
W_CONS   = 20000    # penalty for violating consecutive-day limits

# Wage parameters (€/shift), indexed as [nurse_type][shift_code 0..3]
# nurse_type: 0 = type 1 nurse, 1 = type 2 nurse
WAGE_WEEKDAY = {
    0: [160.0, 160.0, 176.0, 216.0],   # type 1: E, D, L, N
    1: [120.0, 120.0, 132.0, 162.0],   # type 2: E, D, L, N
}

WAGE_WEEKEND = {
    0: [216.0, 216.0, 237.6, 291.6],   # type 1: E, D, L, N
    1: [162.0, 162.0, 178.2, 218.7],   # type 2: E, D, L, N
}

# CONSTANTS 
NURSES = 32
DAYS = 28
SHIFTS = 5
TYPES = 2

# GENERIC PERSONNEL ROSTERING VARIABLES
department: str = ""         
number_days: int = 0          
number_nurses: int = 0        
number_shifts: int = 0        
shift_code: int = 0           


# SHIFT SYSTEM

hrs = [0 for _ in range(SHIFTS)]
req = [[0 for _ in range(SHIFTS)] for _ in range(DAYS)]
shift = [0 for _ in range(SHIFTS)]
start_shift = [0 for _ in range(SHIFTS)]
end_shift = [0 for _ in range(SHIFTS)]
length: int = 0


# PERSONNEL CHARACTERISTICS
number_types: int = 0                 
nurse_type = [0 for _ in range(NURSES)]
pref = [[[0 for _ in range(SHIFTS)] for _ in range(DAYS)]for _ in range(NURSES)]
nurse_percent_employment = [0.0 for _ in range(NURSES)]
personnel_number = ["" for _ in range(NURSES)]


# PERSONNEL ROSTER
cyclic_roster = [[0 for _ in range(DAYS)] for _ in range(NURSES)]
monthly_roster = [[0 for _ in range(DAYS)] for _ in range(NURSES)]


# MONTHLY ROSTER RULES
min_ass = [0 for _ in range(NURSES)]
max_ass = [0 for _ in range(NURSES)]
weekend: int = 0  
identical = [0 for _ in range(NURSES)]

max_cons = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]
min_cons = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]
min_shift = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]
max_shift = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]

min_cons_wrk = [0 for _ in range(NURSES)]
max_cons_wrk = [0 for _ in range(NURSES)]
extreme_max_cons = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]
extreme_min_cons = [[0 for _ in range(SHIFTS)] for _ in range(NURSES)]
extreme_max_cons_wrk: int = 0
extreme_min_cons_wrk: int = 0

# EVALUATION VARIABLES
count_ass: int = 0
count_cons_wrk: int = 0
count_cons: int = 0
count_shift = [0 for _ in range(SHIFTS)]

scheduled = [
    [[0 for _ in range(SHIFTS)] for _ in range(DAYS)]
    for _ in range(TYPES)
]

violations = [0 for _ in range(DAYS * SHIFTS)]

# Generic helpers

def _find_cell_containing(df, text):
    """ Return (row, col) of the first cell whose string contains 'text' (case-insensitive). """
    target = text.upper()
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and target in val.strip().upper():
                return r, c
    raise ValueError(f"Cell containing '{text}' not found in shift system sheet")

def _find_row_starting_with(df, text):
    """Return row index where column 0 starts with 'text' (case-insensitive)."""
    target = text.upper()
    for r in range(df.shape[0]):
        val = df.iat[r, 0]
        if isinstance(val, str) and val.strip().upper().startswith(target):
            return r
    raise ValueError(f"Row with label starting '{text}' not found in constraints sheet")

def _find_col_with_label(df, row_idx, label):
    """Return column index in row 'row_idx' whose cell matches 'label' (case-insensitive)."""
    target = label.upper()
    for c in range(df.shape[1]):
        val = df.iat[row_idx, c]
        if isinstance(val, str) and val.strip().upper().startswith(target):
            return c
    raise ValueError(f"Column with label '{label}' not found in row {row_idx}")

def debug_list_sheets():
    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    xls = pd.ExcelFile(excel_file)
    print("Excel file:", excel_file)
    print("Sheets:", xls.sheet_names)

def is_weekend(day_idx: int) -> bool:
    """ Return True if day_idx (0-based) is Saturday or Sunday."""
    d1 = day_idx + 1  # convert to 1-based
    return (d1 % 7 == 6) or (d1 % 7 == 0)

def max_allowed_workblock(n: int) -> int:
    emp = nurse_percent_employment[n]
    if emp >= 0.99:
        return 5
    elif emp >= 0.74:
        return 4
    return 3

def find_long_workblock(roster, n: int):
    if n >= len(roster):
        raise IndexError(f"Nurse index {n} out of bounds for roster")

    max_len = max_allowed_workblock(n)
    cons = 0
    start = None

    for d in range(number_days + 1):
        if d < number_days and roster[n][d] < FREE_SHIFT:
            cons += 1
            if cons == 1:
                start = d
        else:
            if cons > max_len:
                return start, cons
            cons = 0
            start = None

    return None

# Input reading
def read_shift_system():
    """ Read the shift system for department C from the 'Case_C_9' sheet.
    Internal encoding:
      0 = Early (E)   3 <= start < 9
      1 = Day   (D)   9 <= start < 12
      2 = Late  (L)   12 <= start < 21
      3 = Night (N)   start >= 21 or start < 3
      4 = Free  (F)
    """
    global number_shifts, length, hrs, req, shift, start_shift, end_shift, number_days
    sheet_name = "Case_C_9"   
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    number_shifts = int(df.iat[1, 0])  # A2
    length = int(df.iat[1, 1])         # B2

    r_start, c_start = _find_cell_containing(df, "START SHIFTS DEP C")
    r_req, c_req = _find_cell_containing(df, "REQUIREMENTS DEP C")

    start_rows = [r_start + 1 + i for i in range(number_shifts)]
    req_rows = [r_req + 1 + i for i in range(number_shifts)]

    for j in range(SHIFTS):  # 0..4
        req[0][j] = 0

    # Process each real shift (E/D/L/N)
    for idx in range(number_shifts):
        row_start = start_rows[idx]
        row_req = req_rows[idx]

        start_h = int(df.iat[row_start, c_start])
        required = int(df.iat[row_req, c_req])

        # 1-based index for compatibility with old arrays
        k = idx + 1
        start_shift[k] = start_h

        if 3 <= start_h < 9:
            code = 0  # Early
        elif 9 <= start_h < 12:
            code = 1  # Day
        elif 12 <= start_h < 21:
            code = 2  # Late
        else:
            code = 3  # Night

        shift[k] = code
        hrs[code] = length
        req[0][code] = required

        end_shift[k] = start_h + length if start_h + length < 24 else start_h + length - 24

    # Free shift 
    shift[0] = 4
    hrs[4] = 0 

    # Copy requirements to all days 
    for day in range(1, number_days):
        for j in range(SHIFTS):
            req[day][j] = req[0][j]

    number_shifts = SHIFTS

def read_personnel_characteristics():
    """
    Read nurse preferences, employment, and type from Excel instead of txt.

    E Sheet: 'Case_E_Preferences_<department>'
      col0: Personnel Number (str)
      next 5 * number_days: prefs (day/shift flattened)
      next: employment (float)
      next: type (1 or 2 -> stored as 0 or 1)
    """
    global number_types, personnel_number, pref, nurse_percent_employment, nurse_type, number_nurses, number_days

    sheet_name = f"Case_E_Preferences_{department}"
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    n_prefs = len(df)

    if number_nurses != 0 and n_prefs != number_nurses:
        raise ValueError(
        f"Inconsistent nurse count: cyclic roster has {number_nurses}, "
        f"preferences have {n_prefs}."
    )

    number_nurses = n_prefs
    number_types = TYPES
    
    prefs_per_nurse = SHIFTS * number_days

    for k in range(number_nurses):
        row = df.iloc[k]
        personnel_number[k] = str(row.iloc[0])

        pref_values = row.iloc[1 : 1 + prefs_per_nurse].tolist()

        if len(pref_values) != prefs_per_nurse:
            raise ValueError(
                f"Row {k} in Personnel sheet has {len(pref_values)} preference values, "
                f"expected {prefs_per_nurse}."
            )

        idx = 0
        for day in range(number_days):
            for s in range(SHIFTS):  
                pref[k][day][s] = int(pref_values[idx])
                idx += 1

        employment_col = 1 + prefs_per_nurse
        type_col = employment_col + 1

        nurse_percent_employment[k] = float(row.iloc[employment_col])
        nurse_type[k] = int(row.iloc[type_col]) - 1  


def read_cyclic_roster():
    """
    Read the cyclic roster for this department.

    Sheet: 'Case_D_Cyclic_<department>'
      columns: NurseType, Day1, Day2, ..., DayN
      NurseType: 1,2,... -> stored as 0,1,...
      Day*: shift codes in internal encoding (0..4)
    """

    # Mapping van Excel naar interne shiftcode
    EXCEL_TO_INTERNAL_SHIFT = {
        0: 4,  # Free
        1: 0,  # Early
        2: 1,  # Day
        3: 2,  # Late
        4: 3   # Night
    }


    global number_nurses, number_days, cyclic_roster, nurse_type

    sheet_name = f"Case_D_Cyclic_{department}"
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    n_cyc = len(df)

    if number_nurses != 0 and n_cyc != number_nurses:
        raise ValueError(
            f"Inconsistent nurse count: previous input had {number_nurses}, "
            f"cyclic roster has {n_cyc}."
        )

    number_nurses = n_cyc

    if "NurseType" not in df.columns:
        raise ValueError(f"'NurseType' column missing in sheet {sheet_name} of {excel_file}")

    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(f"No Day* columns found in sheet {sheet_name} of {excel_file}")

    excel_days = len(day_cols)
    if number_days != excel_days:
        print(
            f"WARNING: number_days in code = {number_days}, "
            f"but Excel has {excel_days} day columns. Using {excel_days}."
        )
        number_days = excel_days

    for k in range(number_nurses):
        nt_val = int(df.iloc[k]["NurseType"])
        nurse_type[k] = nt_val - 1
        for d_idx, col in enumerate(day_cols):
            excel_code = int(df.iloc[k][col])
            if excel_code not in EXCEL_TO_INTERNAL_SHIFT:
                raise ValueError(f"Unknown shift code {excel_code} in row {k+1}, column {col}")
            cyclic_roster[k][d_idx] = EXCEL_TO_INTERNAL_SHIFT[excel_code]


def read_monthly_roster_constraints():
    """
    Read monthly roster constraints (Case_E_Constraints_A).

    Fills: min_ass, max_ass, min_cons_wrk, max_cons_wrk,
           min_cons, max_cons, min_shift, max_shift, identical, etc.
    """
    global min_ass, max_ass, min_cons_wrk, max_cons_wrk
    global min_cons, max_cons, extreme_max_cons, extreme_min_cons
    global min_shift, max_shift, identical
    global extreme_max_cons_wrk, extreme_min_cons_wrk

    sheet_name = "Case_E_Constraints_A"
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Total assignments
    r_ass = _find_row_starting_with(df, "NUMBER OF ASSIGNMENTS")
    header_row_ass = r_ass + 1
    val_row_ass = header_row_ass + 1

    min_col = _find_col_with_label(df, header_row_ass, "Minimum")
    max_col = _find_col_with_label(df, header_row_ass, "Maximum")

    base_min_ass = int(df.iat[val_row_ass, min_col])
    base_max_ass = int(df.iat[val_row_ass, max_col])

    # Global consecutive assignments
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

    # Consecutive assignments per shift type
    r_cons_sh = _find_row_starting_with(df, "NUMBER OF CONSECUTIVE ASSIGNMENTS PER SHIFT TYPE")
    header_row_cons_sh = r_cons_sh + 1
    first_val_row_cons_sh = header_row_cons_sh + 1

    min_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Minimum")
    max_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Maximum")

    base_min_cons = {}
    base_max_cons = {}
    sh = 0
    r = first_val_row_cons_sh
    while r < df.shape[0]:
        val_min = df.iat[r, min_col_cons_sh]
        val_max = df.iat[r, max_col_cons_sh]
        if (isinstance(val_min, (float, int))) and not pd.isna(val_min):
            base_min_cons[sh] = int(val_min)
            base_max_cons[sh] = int(val_max)
            sh += 1
            r += 1
        else:
            break
    num_working_shifts = sh  # e.g. 3 (E,D,L)

    # Assignments per shift type
    r_ass_sh = _find_row_starting_with(df, "NUMBER OF ASSIGNMENTS PER SHIFT TYPE")
    header_row_ass_sh = r_ass_sh + 1
    first_val_row_ass_sh = header_row_ass_sh + 1

    min_col_ass_sh = _find_col_with_label(df, header_row_ass_sh, "Minimum")
    max_col_ass_sh = _find_col_with_label(df, header_row_ass_sh, "Maximum")

    base_min_shift = {}
    base_max_shift = {}
    sh = 0
    r = first_val_row_ass_sh
    while r < df.shape[0] and sh < num_working_shifts:
        val_min = df.iat[r, min_col_ass_sh]
        val_max = df.iat[r, max_col_ass_sh]
        if (isinstance(val_min, (float, int))) and not pd.isna(val_min):
            base_min_shift[sh] = int(val_min)
            base_max_shift[sh] = int(val_max)
            sh += 1
            r += 1
        else:
            break

    # Identical weekend
    r_ident = _find_row_starting_with(df, "IDENTICAL WEEKEND CONSTRAINT")
    val_row_ident = r_ident + 1

    ident_value = None
    for c in range(df.shape[1]):
        cell = df.iat[val_row_ident, c]
        if isinstance(cell, str) and cell.strip():
            ident_value = cell.strip().upper()
            break
    ident_flag = 1 if (ident_value and ident_value.startswith("Y")) else 0

    # Apply to all nurses
    for k in range(number_nurses):
        min_ass[k] = int(base_min_ass * nurse_percent_employment[k])
        max_ass[k] = int(base_max_ass * nurse_percent_employment[k])

        min_cons_wrk[k] = base_min_cons_wrk
        max_cons_wrk[k] = base_max_cons_wrk
        extreme_max_cons_wrk = 10
        extreme_min_cons_wrk = 1

        for sh in range(num_working_shifts):
            min_cons[k][sh] = base_min_cons[sh]
            max_cons[k][sh] = base_max_cons[sh]
            extreme_max_cons[k][sh] = 10
            extreme_min_cons[k][sh] = 1
            min_shift[k][sh] = base_min_shift[sh]
            max_shift[k][sh] = base_max_shift[sh]

        for sh in range(num_working_shifts, SHIFTS):

    # If this is the FREE shift (code 4), we don't care about consecutive days
            if sh == 4:
                min_cons[k][sh] = 0
                max_cons[k][sh] = 9999   # any length is fine
            else:
                # Real shift (e.g. Night) but not explicitly in the Excel block:
                # give it a very large max so it almost never violates
                min_cons[k][sh] = 0
                max_cons[k][sh] = 28     # or base_max_cons_wrk, or similar

            extreme_max_cons[k][sh] = 10
            extreme_min_cons[k][sh] = 1
            min_shift[k][sh] = 0
            max_shift[k][sh] = 9999


        identical[k] = ident_flag



def read_monthly_roster_from_excel():
    """
    Read the monthly roster from Excel and fill monthly_roster.

    Sheet: 'Case_E_MonthlyRoster_<department>'
      columns: Personnel Number (optional), Day1..DayN
      Day*: shift codes (0..4)
    """
    # Mapping Excel → interne shiftcode
    EXCEL_TO_INTERNAL_SHIFT = {
        0: 4,  # Free
        1: 0,  # Early
        2: 1,  # Day
        3: 2,  # Late
        4: 3   # Night
    }


    global monthly_roster, number_nurses, number_days

    sheet_name = f"Case_E_MonthlyRoster_{department}"
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    if "Personnel Number" in df.columns:
        for k in range(min(number_nurses, len(df))):
            roster_id = str(df.iloc[k]["Personnel Number"])
            if roster_id != personnel_number[k]:
                raise ValueError(
                    f"Mismatch between preferences and monthly roster at row {k}: "
                    f"prefs PN = {personnel_number[k]}, roster PN = {roster_id}"
                )

    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(f"No Day* columns found in sheet {sheet_name} of {excel_file}")

    excel_days = len(day_cols)
    if excel_days != number_days:
        print(
            f"WARNING: code expects {number_days} days, "
            f"but Excel monthly roster has {excel_days} days. Using {excel_days}."
        )
        number_days = excel_days

    if len(df) < number_nurses:
        raise ValueError(
            f"Monthly roster has only {len(df)} nurses, "
            f"but number_nurses = {number_nurses} from other input."
        )
    if len(df) > number_nurses:
        print(
            f"WARNING: Monthly roster has {len(df)} rows but number_nurses = {number_nurses}. "
            f"Ignoring extra rows."
        )

    for k in range(number_nurses):
        for d_idx, col in enumerate(day_cols):
            excel_code = int(df.iloc[k][col])
            if excel_code not in EXCEL_TO_INTERNAL_SHIFT:
                raise ValueError(f"Unknown shift code {excel_code} for nurse {k+1}, day {d_idx+1}")
            monthly_roster[k][d_idx] = EXCEL_TO_INTERNAL_SHIFT[excel_code]

def debug_print_first_nurse():
    print("\n=== QUICK CHECK: FIRST NURSE DATA ===")

    # ID & FTE
    print(f"Personnel Number: {personnel_number[0]}")
    print(f"Employment %: {nurse_percent_employment[0]:.2f}")
    print(f"Type: {nurse_type[0]+1}")  # +1 to match Excel (1/2)

    # Preferences
    print("Preferences (Day 0):")
    for s in range(SHIFTS):
        label = SHIFT_LABELS.get(s, f"Shift{s}")
        print(f"  {label}: {pref[0][0][s]}")

    # Assignment constraints
    print(f"Assignment Min: {min_ass[0]}, Max: {max_ass[0]}")
    print(f"Consecutive Working Days: min = {min_cons_wrk[0]}, max = {max_cons_wrk[0]}")

    print("Consecutive Shifts Per Type:")
    for s in range(SHIFTS):
        label = SHIFT_LABELS.get(s, f"Shift{s}")
        print(f"  {label}: min = {min_cons[0][s]}, max = {max_cons[0][s]}")

    print("Assignment Limits Per Shift:")
    for s in range(SHIFTS):
        label = SHIFT_LABELS.get(s, f"Shift{s}")
        print(f"  {label}: min = {min_shift[0][s]}, max = {max_shift[0][s]}")

    print(f"Identical Weekend Constraint: {'YES' if identical[0] else 'NO'}")

    print("======================================\n")

def debug_capacity_vs_demand():
    total_required = 0
    for d in range(number_days):
        for s in range(number_shifts - 1):  # ignore F
            total_required += req[d][s]

    total_max_assign = sum(max_ass)
    total_min_assign = sum(min_ass)
    print("Total required assignments:", total_required)
    print("Total min assignments:", total_min_assign)
    print("Total max assignments:", total_max_assign)
    if total_required > total_max_assign:
        print(">> INFEASIBLE: demand exceeds total max assignments.")
    elif total_required < total_min_assign:
        print(">> OVERCAPACITY: even min_ass exceeds demand.")
    else:
        print(">> Globally feasible in terms of total capacity.")


def read_input():
    """Read all input and initialise data structures."""
    global number_shifts
    read_shift_system()
    read_cyclic_roster()
    read_personnel_characteristics()
    read_monthly_roster_constraints()
    
    # --- Enforce contract-based minimum shifts as HARDER lower bound ---
    for n in range(number_nurses):
        emp = nurse_percent_employment[n]

        if emp >= 0.99:          # full time
            contract_min = 20
        elif emp >= 0.74:        # 0.75 FTE
            contract_min = 15
        else:                    # optional fallback (e.g. 0.5 FTE -> 10)
            contract_min = 10

        # Take the MAX between Excel min and contract min
        min_ass[n] = max(min_ass[n], contract_min)
    
    number_shifts = SHIFTS


def print_output():
    """
    Print the monthly roster to txt and return DataFrame (labels E/D/L/N/F).
    """
    txt_filename = f"Monthly_Roster_dpt_{department}.txt"

    with open(txt_filename, "w") as f:
        for k in range(number_nurses):
            f.write(f"{personnel_number[k]}\t")
            for i in range(number_days):
                code = monthly_roster[k][i]
                f.write(f"{code}\t")
            f.write("\n")
    print(f"Monthly roster written to {txt_filename}")

    data = {"Personnel Number": [personnel_number[k] for k in range(number_nurses)]}
    for d in range(number_days):
        colname = f"Day{d + 1}"
        col = []
        for k in range(number_nurses):
            code = monthly_roster[k][d]
            col.append(code)
        data[colname] = col

    return pd.DataFrame(data)



def evaluate_line_of_work(nurse_idx: int, slack_j: int = 0):
    """
    Evaluate the monthly roster line for nurse `nurse_idx`.

    Updates:
      - violations[0..4]
      - scheduled[type][day][shift]
    """
    global count_ass, count_cons_wrk, count_cons, count_shift

    i = nurse_idx
    j = slack_j

    hh = 0
    count_ass = 0
    count_cons_wrk = 0
    count_cons = 0
    for l in range(number_shifts):
        count_shift[l] = 0

    # Day 0
    a = monthly_roster[i][0]
    violations[0] += pref[i][0][a]

    if a < 4:
        count_ass += 1
        count_cons_wrk += 1
        count_cons += 1

    count_shift[a] += 1
    kk = nurse_type[i]
    scheduled[kk][0][a] += 1

    # Remaining days
    for k in range(1, number_days):
        h1 = monthly_roster[i][k]
        h2 = monthly_roster[i][k - 1]

        scheduled[kk][k][h1] += 1
        violations[0] += pref[i][k][h1]

        if h1 < 4:
            count_ass += 1
        count_shift[h1] += 1

        if h1 < 4:
            count_cons_wrk += 1
        elif h1 == 4 and h2 < 4:
            if count_cons_wrk > max_cons_wrk[i] + j:
                violations[1] += 1
            count_cons_wrk = 0

        if h1 != h2:
            if count_cons > max_cons[i][h2] + j:
                violations[2] += 1
            count_cons = 1
        else:
            count_cons += 1

    if count_ass < min_ass[i]:
        violations[3] += 1
    if count_ass > max_ass[i]:
        violations[4] += 1


def evaluate_solution():
    """
    Evaluate the current monthly_roster.

    Returns:
      df_summary, df_staffing
    """
    for kk in range(number_types):
        for day in range(number_days):
            for sh in range(number_shifts):
                scheduled[kk][day][sh] = 0

    for idx in range(20):
        violations[idx] = 0

    for nurse_idx in range(number_nurses):
        evaluate_line_of_work(nurse_idx)

    txt_filename = BASE_DIR / f"Violations_dpt_{department}.txt"
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
        for day in range(number_days):
            for sh in range(number_shifts - 1):  # ignore free shift
                total_scheduled = sum(scheduled[kk][day][sh] for kk in range(number_types))
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
    for day in range(number_days):
        for sh in range(number_shifts - 1):  # ignore free shift
            total_scheduled = sum(scheduled[kk][day][sh] for kk in range(number_types))
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

def count_consecutive_shifttype_violations(roster):
    """
    Count how many times the 'max_cons per shift type' rule is violated,
    in the same spirit as evaluate_line_of_work, but for an arbitrary roster.
    """
    total_viol = 0

    for i in range(number_nurses):
        count_cons = 0
        prev_s = roster[i][0]

        for k in range(1, number_days + 1):
            if k < number_days:
                s = roster[i][k]
            else:
                s = -1  # force closure at end

            if s != prev_s:
                # end of block of prev_s
                if prev_s < 4 and count_cons > max_cons[i][prev_s]:
                    total_viol += (count_cons - max_cons[i][prev_s])
                count_cons = 1 if (s < 4) else 0
            else:
                if s < 4:
                    count_cons += 1

            prev_s = s

    return total_viol


# Objective 

WEIGHT_WAGE     = 1
WEIGHT_NURSE    = 1
WEIGHT_PATIENT  = 1


def compute_components(roster):
    """
    Compute (wage_cost, nurse_cost, patient_cost) for a given roster.

    roster: 2D list [nurse][day] with shift codes 0..4.
    """
    wage_cost = 0.0
    nurse_cost = 0.0
    patient_cost = 0.0
    
    PREF_PEN = 1          # already defined below, but better to keep together


    # 1) Wage cost
    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            continue

        t = nurse_type[n]  # 0 = type 1, 1 = type 2
        for d in range(number_days):
            s = roster[n][d]
            if s < 4:  # E/D/L/N only
                weekend_flag = is_weekend(d)
                if weekend_flag:
                    wage_cost += WAGE_WEEKEND[t][s]
                else:
                    wage_cost += WAGE_WEEKDAY[t][s]

    # 2) Patient satisfaction (as cost)
    SHIFT_CHANGE_PEN = 1.0
    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            continue

        for d in range(1, number_days):
            s_prev = roster[n][d - 1]
            s_curr = roster[n][d]
            if s_prev < 4 and s_curr < 4 and s_prev != s_curr:
                patient_cost += SHIFT_CHANGE_PEN

    for d in range(number_days):
        for s in range(number_shifts - 1):
            scheduled_count = sum(roster[n][d] == s for n in range(number_nurses))
            diff = scheduled_count - req[d][s]
            if diff < 0:
                shortage = -diff
                patient_cost += W_UNDER * (shortage ** 2)
            elif diff > 0:
                surplus = diff
                patient_cost += W_OVER * (surplus ** 2)


    # 3) Nurse satisfaction (as cost)
    LATE_SHIFT = 2
    EARLY_SHIFT = 0
    NIGHT_SHIFT = 3
    


    LATE_EARLY_PEN = 1
    NIGHT_REST_PEN = 1
    CONTRACT_MIN_PEN = 1  # penalty per missing shift vs contract minimum
    CONS_WORK_PEN = 1
    
    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            continue

        # Succession
        for d in range(1, number_days):
            prev_s = roster[n][d - 1]
            curr_s = roster[n][d]

            if prev_s == LATE_SHIFT and curr_s == EARLY_SHIFT:
                nurse_cost += LATE_EARLY_PEN

            if prev_s == NIGHT_SHIFT and curr_s in (EARLY_SHIFT, LATE_SHIFT):
                nurse_cost += NIGHT_REST_PEN

        # Max consecutive working days – use max_cons_wrk[n] from constraints
        limit = max_cons_wrk[n]
        if limit > 0:
            cons = 0
            for d in range(number_days + 1):
                if d < number_days and roster[n][d] < 4:
                    cons += 1
                else:
                    if cons > limit:
                        nurse_cost += CONS_WORK_PEN * (cons - limit)
                    cons = 0

        # --- Contract-based minimum shifts (full-time vs part-time) ---
        worked = sum(roster[n][d] < 4 for d in range(number_days))
        emp = nurse_percent_employment[n]  # e.g. 1.0, 0.75, 0.5 ...

        # classify contract:
        #   >= 0.99  -> full-time  -> at least 20 shifts
        #   else     -> part-time -> at least 15 shifts
        if emp >= 0.99:
            min_contract_shifts = 20
        else:
            min_contract_shifts = 15

        # big penalty for being below the contract minimum
        if worked < min_contract_shifts:
            nurse_cost += CONTRACT_MIN_PEN * (min_contract_shifts - worked)

        # Hard minimum from Excel (usually scaled with employment)
        if worked < min_ass[n]:
            nurse_cost += CONTRACT_MIN_PEN * (min_ass[n] - worked)

        # optional: still discourage crazy overscheduling using max_ass
        if worked > max_ass[n]:
            nurse_cost += W_ASSIGN * (worked - max_ass[n])

        
        # Individual preferences
        for d in range(number_days):
            s = roster[n][d]
            if s < 4:
                nurse_cost += PREF_PEN * pref[n][d][s]

    # 4) Extra penalty for per-shift-type consecutive violations (once per roster)
    cons_shift_viol = count_consecutive_shifttype_violations(roster)
    nurse_cost += W_CONS * cons_shift_viol

    return wage_cost, nurse_cost, patient_cost

    


def violates_contract_min(roster) -> bool:
    """
    Return True if ANY nurse violates the hard contract minimum:
      - full-time (emp >= 0.99): at least 20 shifts
      - part-time (emp < 0.99): at least 15 shifts
    """
    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            # if they never work in this department, treat as violation or not?
            # Here we treat it as violation because of the contract.
            return True

        emp = nurse_percent_employment[n]
        worked = sum(roster[n][d] < 4 for d in range(number_days))

        if emp >= 0.99:
            min_contract_shifts = 20
        else:
            min_contract_shifts = 15

        if worked < min_contract_shifts:
            return True

    return False


def random_neighbor(roster, p_swap=0.4, p_fix_block=0.3):
    new_roster = deepcopy(roster)

    if number_nurses < 1 or number_days < 1:
        return new_roster

    #  Eerst: probeer een te lang werkblok te breken
    for n in random.sample(range(number_nurses), k=number_nurses):
        block = find_long_workblock(new_roster, n)
        if block is not None:
            start, length = block
            # kies een dag binnen dat te lange blok om vrij te maken
            d = random.randint(start, start + length - 1)
            new_roster[n][d] = 4  # 4 = Free
            return new_roster

    #  Pas als er geen illegale werkblokken zijn: je oude moves
    r = random.random()

    if r < p_swap:
        # swap move (as before)
        ...
    elif r < p_swap + p_fix_block:
        
        # Search for any nurse and shift type with a violating block
        violation_found = False
        for n in random.sample(range(number_nurses), k=number_nurses):
            for s in range(number_shifts - 1):  # iterate over each actual shift type (0..3)
                limit = max_cons[n][s]
                if limit <= 0:
                    continue  # skip if no limit defined
                cons = 0
                start = None
                for d in range(number_days + 1):  # include end sentinel
                    if d < number_days and new_roster[n][d] == s:
                        # building a block of shift s
                        cons += 1
                        if cons == 1:
                            start = d
                    else:
                        # block ended
                        if cons > limit:
                            # Found a violation: break the block
                            change_day = random.randint(start, start + cons - 1)
                            # Change the shift on change_day to something else (preferably a non-violating option)
                            possible_shifts = [x for x in range(number_shifts) if x != s]
                            if possible_shifts:
                                new_roster[n][change_day] = random.choice(possible_shifts)
                            violation_found = True
                            break
                        cons = 0
                if violation_found:
                    break
            if violation_found:
                break

        if not violation_found:
            # No consecutive-shift violation was found; fall back to a simple random change
            n = random.randrange(number_nurses)
            d = random.randrange(number_days)
            old_shift = new_roster[n][d]
            possible_shifts = [x for x in range(number_shifts) if x != old_shift]
            if possible_shifts:
                new_roster[n][d] = random.choice(possible_shifts)

    else:
        # simple change-coverage move
        ...
    return new_roster

def simulated_annealing(initial_roster,
                        T_start=1000.0,
                        T_min=1e-3,
                        alpha=0.95,
                        iters_per_T=1000):
    current = deepcopy(initial_roster)
    best = deepcopy(initial_roster)
    current_cost = compute_objective(current)
    best_cost = current_cost
    T = T_start

    while T > T_min:
        for _ in range(iters_per_T):

            # --- draw a FEASIBLE neighbour (w.r.t. contract mins) ---
            attempts = 0
            neighbor = None
            # --- Enhanced neighbor selection with consecutive constraint check ---
            while attempts < 30:
                candidate = random_neighbor(current)
                # Skip candidate if it breaks any hard constraints:
                if violates_contract_min(candidate):
                    attempts += 1
                    continue
                # Also ensure we don't increase consecutive shift-type violations
                curr_viol = count_consecutive_shifttype_violations(current)
                cand_viol = count_consecutive_shifttype_violations(candidate)
                if cand_viol > curr_viol:
                    # Candidate makes consecutive-shift violations worse – skip it
                    attempts += 1
                    continue
                # Otherwise, accept this neighbor for cost evaluation
                neighbor = candidate
                break


            if neighbor is None:
                # couldn't find a feasible neighbour this iteration
                continue

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


def procedure():
    """
    Construct and improve the monthly roster for `department` using SA.
    """
    global monthly_roster

    read_monthly_roster_from_excel()

    initial_roster = [
        [monthly_roster[n][d] for d in range(number_days)]
        for n in range(number_nurses)
    ]

    w0, n0, p0 = compute_components(initial_roster)
    obj0 = compute_objective(initial_roster)
    print("Initial schedule metrics:")
    print(f"  Wage_cost      = {w0:.2f}")
    print(f"  Nurse_cost     = {n0:.2f}")
    print(f"  Patient_cost   = {p0:.2f}")
    print(f"  Objective      = {obj0:.2f}")


    best_roster, best_obj = simulated_annealing(
        initial_roster,
        T_start=1000.0,
        T_min=1e-3,
        alpha=0.95,
        iters_per_T=200
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


def add_nurse_to_day_shift(nurse_id: int, day_id: int, shift_id: int):
    """Assign nurse to shift on given day (internal encoding)."""
    monthly_roster[nurse_id][day_id] = shift_id

def compute_objective(roster):
    """Weighted objective."""
    wage_cost, nurse_cost, patient_cost = compute_components(roster)

    return (
        WEIGHT_WAGE * wage_cost +
        WEIGHT_NURSE * nurse_cost +
        WEIGHT_PATIENT * patient_cost
    )


def main():
    global number_days, weekend, department, elapsed_time

    number_days = 28
    weekend = 7
    department = "C"

    seed = 1000
    random.seed(seed)

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

    output_file = BASE_DIR / f"CASE_E_output_{department}.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_roster.to_excel(writer, sheet_name="MonthlyRoster", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_staffing.to_excel(writer, sheet_name="StaffingViolations", index=False)

    print(f"\nExcel output written to: {output_file}")

if __name__ == "__main__":
    main()
