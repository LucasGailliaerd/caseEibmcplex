import math
import time
import random
import pandas as pd
from pathlib import Path
from copy import deepcopy

# Paths and constants

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "CASE_E_input.xlsx"

SHIFT_LABELS = {0: "E", 1: "D", 2: "L", 3: "N", 4: "F"}

# Objective weights
W_PREF   = 1.0      # nurse dissatisfaction (preference score)
W_UNDER  = 1000.0   # penalty per nurse missing (understaffing)
W_OVER   = 100.0    # penalty per nurse extra (overstaffing)
W_ASSIGN = 50.0     # penalty per shifts beyond min/max total assignments
W_CONS   = 50.0     # penalty for violating consecutive-day limits

# Wage parameters (€/shift for example)
WAGE_TYPE1_WEEKDAY = 1.0
WAGE_TYPE1_WEEKEND = 1.5
WAGE_TYPE2_WEEKDAY = 0.8
WAGE_TYPE2_WEEKEND = 1.2


# CONSTANTS 
NURSES = 100
DAYS = 30
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

# Input reading
def read_shift_system():
    """ Read the shift system for department A from the 'Case_C_9' sheet.
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

    r_start, c_start = _find_cell_containing(df, "START SHIFTS DEP A")
    r_req, c_req = _find_cell_containing(df, "REQUIREMENTS DEP A")

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

    Excel file:
      - CASE_E_input.xlsx
      - sheet: 'Personnel_A' (for department 'A')
      - NO HEADER ROW
      - Each row:
          col0: Personnel Number (string)
          col1..col(1+5*number_days-1):  preference ints (flattened: day 1 shift0..4, day 2 shift0..4, ...)
          next col: employment (float, e.g. 1.00)
          next col: type (1 or 2, will be stored as 0 or 1 internally)
    """
    global number_types, personnel_number, pref, nurse_percent_employment, nurse_type, number_nurses, number_days

    excel_file = BASE_DIR / "CASE_E_input.xlsx"
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

    

    # total prefs per nurse = 5 * number_days
    prefs_per_nurse = 5 * number_days

    for k in range(number_nurses):
        row = df.iloc[k]

        # 0: personnel number
        personnel_number[k] = str(row.iloc[0])

        # 1..prefs_per_nurse: flattened preferences
        pref_values = row.iloc[1 : 1 + prefs_per_nurse].tolist()

        if len(pref_values) != prefs_per_nurse:
            raise ValueError(
                f"Row {k} in Personnel sheet has {len(pref_values)} preference values, "
                f"expected {prefs_per_nurse}."
            )

        idx = 0
        for day in range(number_days):
            for s in range(5):  # 5 shift types
                pref[k][day][s] = int(pref_values[idx])
                idx += 1

        # employment and type
        employment_col = 1 + prefs_per_nurse
        type_col = employment_col + 1

        nurse_percent_employment[k] = float(row.iloc[employment_col])
        nurse_type[k] = int(row.iloc[type_col]) - 1  # make it 0 or 1 internally


def read_cyclic_roster():
    """
    Read the cyclic roster for this department from Excel.

    Excel file:
      - CASE_E_input.xlsx  (must be in the same folder as this .py)
      - sheet: 'CyclicRoster_<department>', e.g. 'CyclicRoster_A'
      - columns:
          NurseType, Day1, Day2, ..., DayN
      - NurseType = 1,2,... (will be stored internally as 0,1,...)
      - Day* cells = external shift codes (indices into shift[])
    """
    global number_nurses, number_days, cyclic_roster, nurse_type

    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    sheet_name = f"Case_D_Cyclic_{department}"

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    n_cyc = len(df)

    
    if number_nurses != 0 and n_cyc != number_nurses:
        raise ValueError(
            f"Inconsistent nurse count: previous input had {number_nurses}, "
            f"cyclic roster has {n_cyc}."
        )

    number_nurses = n_cyc


    # Check NurseType column
    if "NurseType" not in df.columns:
        raise ValueError(
            f"'NurseType' column missing in sheet {sheet_name} of {excel_file}"
        )

    # All day columns = those starting with "Day"
    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(
            f"No Day* columns found in sheet {sheet_name} of {excel_file}"
        )

    # Number of days from Excel
    excel_days = len(day_cols)
    if number_days != excel_days:
        print(
            f"WARNING: number_days in code = {number_days}, "
            f"but Excel has {excel_days} day columns. "
            f"Using {excel_days} from Excel."
        )
        number_days = excel_days


    # Fill nurse_type and cyclic_roster
    for k in range(number_nurses):
        nt_val = int(df.iloc[k]["NurseType"])
        nurse_type[k] = nt_val - 1  # type 1/2 -> 0/1

        for d_idx, col in enumerate(day_cols):
            code = int(df.iloc[k][col])      # 0=E,1=D,2=L,3=N,4=F
            cyclic_roster[k][d_idx] = code   # already internal encoding


def read_monthly_roster_constraints():
    """
    Read monthly roster constraints for Case E, Department A from the
    'Case_E_Constraints_A' sheet with layout:

    DEPARTMENT A
    NUMBER OF ASSIGNMENTS
      Minimum  Maximum
      0        20

    NUMBER OF CONSECUTIVE ASSIGNMENTS
      Minimum  Maximum
      1        6

    NUMBER OF CONSECUTIVE ASSIGNMENTS PER SHIFT TYPE
      Minimum  Maximum
      1        6
      1        6
      1        28

    NUMBER OF ASSIGNMENTS PER SHIFT TYPE
      Minimum  Maximum
      0        12
      0        12
      0        12

    IDENTICAL WEEKEND CONSTRAINT
      NO
    """
    global min_ass, max_ass, min_cons_wrk, max_cons_wrk
    global min_cons, max_cons, extreme_max_cons, extreme_min_cons
    global min_shift, max_shift, identical
    global extreme_max_cons_wrk, extreme_min_cons_wrk

    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    sheet_name = "Case_E_Constraints_A"

    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # ----- NUMBER OF ASSIGNMENTS -----
    r_ass = _find_row_starting_with(df, "NUMBER OF ASSIGNMENTS")
    header_row_ass = r_ass + 1
    val_row_ass = header_row_ass + 1

    min_col = _find_col_with_label(df, header_row_ass, "Minimum")
    max_col = _find_col_with_label(df, header_row_ass, "Maximum")

    base_min_ass = int(df.iat[val_row_ass, min_col])
    base_max_ass = int(df.iat[val_row_ass, max_col])

    # ----- NUMBER OF CONSECUTIVE ASSIGNMENTS (GLOBAL) -----
    r_cons = _find_row_starting_with(df, "NUMBER OF CONSECUTIVE ASSIGNMENTS")
    # ensure we don't pick the "PER SHIFT TYPE" block
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

    # ----- CONSECUTIVE ASSIGNMENTS PER SHIFT TYPE -----
    r_cons_sh = _find_row_starting_with(df, "NUMBER OF CONSECUTIVE ASSIGNMENTS PER SHIFT TYPE")
    header_row_cons_sh = r_cons_sh + 1
    first_val_row_cons_sh = header_row_cons_sh + 1

    min_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Minimum")
    max_col_cons_sh = _find_col_with_label(df, header_row_cons_sh, "Maximum")

    base_min_cons = {}
    base_max_cons = {}

    # collect all consecutive non-empty rows in that block
    sh = 0
    r = first_val_row_cons_sh
    while r < df.shape[0]:
        val_min = df.iat[r, min_col_cons_sh]
        val_max = df.iat[r, max_col_cons_sh]
        if (isinstance(val_min, float) or isinstance(val_min, int)) and not pd.isna(val_min):
            base_min_cons[sh] = int(val_min)
            base_max_cons[sh] = int(val_max)
            sh += 1
            r += 1
        else:
            break
    num_working_shifts = sh  # for dept A this will be 3 (E,D,L)

    # ----- ASSIGNMENTS PER SHIFT TYPE -----
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
        if (isinstance(val_min, float) or isinstance(val_min, int)) and not pd.isna(val_min):
            base_min_shift[sh] = int(val_min)
            base_max_shift[sh] = int(val_max)
            sh += 1
            r += 1
        else:
            break

    # ----- IDENTICAL WEEKEND CONSTRAINT -----
    r_ident = _find_row_starting_with(df, "IDENTICAL WEEKEND CONSTRAINT")
    val_row_ident = r_ident + 1

    ident_value = None
    for c in range(df.shape[1]):
        cell = df.iat[val_row_ident, c]
        if isinstance(cell, str) and cell.strip():
            ident_value = cell.strip().upper()
            break
    ident_flag = 1 if (ident_value and ident_value.startswith("Y")) else 0

    # ----- APPLY TO ALL NURSES -----
    for k in range(number_nurses):
        # total assignments (scaled by employment rate)
        min_ass[k] = int(base_min_ass * nurse_percent_employment[k])
        max_ass[k] = int(base_max_ass * nurse_percent_employment[k])

        # global consecutive working days
        min_cons_wrk[k] = base_min_cons_wrk
        max_cons_wrk[k] = base_max_cons_wrk
        extreme_max_cons_wrk = 10
        extreme_min_cons_wrk = 1

        # per-shift rules for the working shifts we actually have (0..num_working_shifts-1)
        for sh in range(num_working_shifts):
            min_cons[k][sh] = base_min_cons[sh]
            max_cons[k][sh] = base_max_cons[sh]
            extreme_max_cons[k][sh] = 10
            extreme_min_cons[k][sh] = 1

            min_shift[k][sh] = base_min_shift[sh]
            max_shift[k][sh] = base_max_shift[sh]

        # for any remaining shifts (night, free, etc.): no requirements
        for sh in range(num_working_shifts, SHIFTS):
            min_cons[k][sh] = 0
            max_cons[k][sh] = 0
            extreme_max_cons[k][sh] = 10
            extreme_min_cons[k][sh] = 1
            min_shift[k][sh] = 0
            max_shift[k][sh] = 9999

        identical[k] = ident_flag



def read_monthly_roster_from_excel():
    """
    Read the monthly roster from Excel and fill monthly_roster.

    Excel:
      - file: CASE_E_input.xlsx (same folder as this .py)
      - sheet: 'MonthlyRoster_<department>', e.g. 'MonthlyRoster_A'
      - columns:
          NurseID, Day1, Day2, ..., DayN
      - Day* cells = external shift codes (same numbering as cyclic input)
    """
    global monthly_roster, number_nurses, number_days

    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    sheet_name = f"Case_E_MonthlyRoster_{department}"

    df = pd.read_excel(excel_file, sheet_name=sheet_name)

 

    # OPTIONAL: check that Personnel Number order matches preferences
    if "Personnel Number" in df.columns:
        for k in range(min(number_nurses, len(df))):
            roster_id = str(df.iloc[k]["Personnel Number"])
            if roster_id != personnel_number[k]:
                raise ValueError(
                    f"Mismatch between preferences and monthly roster at row {k}: "
                    f"prefs PN = {personnel_number[k]}, roster PN = {roster_id}"
                )



    # Identify day columns
    day_cols = [c for c in df.columns if str(c).lower().startswith("day")]
    if not day_cols:
        raise ValueError(
            f"No Day* columns found in sheet {sheet_name} of {excel_file}"
        )

    excel_days = len(day_cols)
    if excel_days != number_days:
        print(
            f"WARNING: code expects {number_days} days, "
            f"but Excel monthly roster has {excel_days} days. "
            f"Using {excel_days} from Excel."
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

    # Fill monthly_roster using internal shift encoding
    for k in range(number_nurses):
        for d_idx, col in enumerate(day_cols):
            code = int(df.iloc[k][col])        # 0..4 = E,D,L,N,F
            monthly_roster[k][d_idx] = code    # store directly


def read_input():
    """
    Read all input files and initialise data structures.
    Equivalent of C++ read_input().
    """
    global cyclic_roster, number_shifts

    read_shift_system()
    read_cyclic_roster()
    read_personnel_characteristics()
    read_monthly_roster_constraints()

    # 4) force number of shifts in algorithm to 5 (E, D, L, N, off)
    number_shifts = 5

def shift_decoding(shift_code: int) -> int:
    """
    Return the index in `shift` that has the given encoded shift_code.
    If not found, return -1.
    """
    for idx in range(number_shifts):
        if shift[idx] == shift_code:
            return idx
    return -1

import os  # at top with the other imports if you want, optional


def print_output():
    """
    Print the monthly roster in the student's original shift numbering.
    Equivalent of C++ print_output().
    """
    txt_filename = f"Monthly_Roster_dpt_{department}.txt"

    with open(txt_filename, "w") as f:
        for k in range(number_nurses):
            f.write(f"{personnel_number[k]}\t")
            for i in range(number_days):
                code = monthly_roster[k][i]  # 0..4 = E,D,L,N,F
                f.write(f"{code}\t")
            f.write("\n")


    print(f"Monthly roster written to {txt_filename}")

    #-----Excel output-----

    data = {}
    data["Personnel Number"] = [personnel_number[k] for k in range(number_nurses)]

    for d in range(number_days):
        colname = f"Day{d+1}"
        col = []
        for k in range(number_nurses):
            code = monthly_roster[k][d]
            col.append(SHIFT_LABELS.get(code, code))
        data[colname] = col

    df = pd.DataFrame(data)
    return df


def evaluate_line_of_work(nurse_idx: int, slack_j: int = 0):
    """
    Evaluate the monthly roster line for nurse `nurse_idx`.

    Updates:
      - violations[0..4]
      - scheduled[type][day][shift]
      - (internally uses count_shift, etc.)
    `slack_j` corresponds to the `+ j` tolerances in the C++ code.
    """
    global count_ass, count_cons_wrk, count_cons, count_shift

    i = nurse_idx
    j = slack_j

    # reset counters
    hh = 0
    count_ass = 0
    count_cons_wrk = 0
    count_cons = 0
    for l in range(number_shifts):
        count_shift[l] = 0

    # day 0
    a = monthly_roster[i][0]

    # preference cost
    violations[0] += pref[i][0][a]

    # working day? (0..3 = work, 4 = day off)
    if a < 4:
        count_ass += 1
        count_cons_wrk += 1
        count_cons += 1

    count_shift[a] += 1
    kk = nurse_type[i]
    scheduled[kk][0][a] += 1

    # remaining days
    for k in range(1, number_days):
        h1 = monthly_roster[i][k]
        h2 = monthly_roster[i][k - 1]

        # record schedule
        scheduled[kk][k][h1] += 1

        # add preference cost
        violations[0] += pref[i][k][h1]

        # min/max assignments (total working days)
        if h1 < 4:
            count_ass += 1

        count_shift[h1] += 1

        # consecutive working days
        if h1 < 4:
            count_cons_wrk += 1
        elif h1 == 4 and h2 < 4:
            # just ended a block of consecutive work days
            if count_cons_wrk > max_cons_wrk[i] + j:
                violations[1] += 1
            count_cons_wrk = 0

        # consecutive same-shift days
        if h1 != h2:
            # ended block of same shift type
            if count_cons > max_cons[i][h2] + j:
                violations[2] += 1
            count_cons = 1  # start new block with h1
        else:
            count_cons += 1

    # after last day: check min/max assignments
    if count_ass < min_ass[i]:
        violations[3] += 1
    if count_ass > max_ass[i]:
        violations[4] += 1


def evaluate_solution():
    """
    Evaluate the current monthly_roster:

      - Reset counters
      - Call evaluate_line_of_work() for each nurse
      - Write txt violations report
      - Return (df_summary, df_staffing) for Excel output
    """
    # Reset scheduled and violations
    for kk in range(number_types):
        for day in range(number_days):
            for sh in range(number_shifts):
                scheduled[kk][day][sh] = 0

    for idx in range(20):
        violations[idx] = 0

    # Evaluate each nurse (uses global i inside evaluate_line_of_work)
    for nurse_idx in range(number_nurses):
        evaluate_line_of_work(nurse_idx)


    # ---------- TXT OUTPUT ----------
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
                        f"There are too few nurses in shift {sh} on day {day+1}: "
                        f"{total_scheduled} < {required}.\n"
                    )
                elif total_scheduled > required:
                    f.write(
                        f"There are too many nurses in shift {sh} on day {day+1}: "
                        f"{total_scheduled} > {required}.\n"
                    )

    print(f"Violations txt written to {txt_filename}")

    # ---------- SUMMARY DATAFRAME ----------
    df_summary = pd.DataFrame([{
        "TotalPreferenceScore": violations[0],
        "MaxConsWorkViol": violations[1],
        "MaxConsShiftViol": violations[2],
        "MinAssignViol": violations[3],
        "MaxAssignViol": violations[4],
    }])

    # ---------- STAFFING VIOLATIONS DATAFRAME ----------
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

# Weights from the Algo.txt
WEIGHT_WAGE     = 0.2
WEIGHT_NURSE    = 10.0
WEIGHT_PATIENT  = 2.0


def compute_components(roster):
    """
    Compute (wage_cost, nurse_cost, patient_cost) for a given roster.

    roster: 2D list [nurse][day] with shift codes 0..4 (E,D,L,N,F).

    Returns:
        wage_cost, nurse_cost, patient_cost
    """

    wage_cost = 0.0
    nurse_cost = 0.0
    patient_cost = 0.0

    # ---------------- 1) Wage cost ----------------
    # For each worked shift (0..3) we add type/weekday/weekend cost.
    for n in range(number_nurses):
        # detect if nurse is actually scheduled in the department (not all off)
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            # "when a nurse is assigned to a line of work containing only zeros,
            # the nurse does not work in the department and penalties are not calculated"
            continue

        t = nurse_type[n]   # 0 or 1

        for d in range(number_days):
            s = roster[n][d]
            if s < 4:  # working day
                weekend_flag = is_weekend(d)

                if t == 0:  # type 1
                    if weekend_flag:
                        wage_cost += WAGE_TYPE1_WEEKEND
                    else:
                        wage_cost += WAGE_TYPE1_WEEKDAY
                else:       # type 2
                    if weekend_flag:
                        wage_cost += WAGE_TYPE2_WEEKEND
                    else:
                        wage_cost += WAGE_TYPE2_WEEKDAY

    # ---------------- 2) Patient satisfaction (as cost) ----------------
    # (a) Penalty when nurses change shifts a lot  -> continuity within line of work
    SHIFT_CHANGE_PEN = 1.0   # tune
    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            continue

        for d in range(1, number_days):
            s_prev = roster[n][d - 1]
            s_curr = roster[n][d]
            # Only count changes between working shifts (exclude free)
            if s_prev < 4 and s_curr < 4 and s_prev != s_curr:
                patient_cost += SHIFT_CHANGE_PEN

    # (b) Penalty when requirements are not met (compare with req[d][shift])
    UNDER_PEN = 1000.0  # very high
    OVER_PEN  = 100.0   # smaller

    for d in range(number_days):
        for s in range(number_shifts - 1):  # 0..3 E,D,L,N
            scheduled = 0
            for n in range(number_nurses):
                if roster[n][d] == s:
                    scheduled += 1
            diff = scheduled - req[d][s]
            if diff < 0:
                patient_cost += UNDER_PEN * (-diff)
            elif diff > 0:
                patient_cost += OVER_PEN * diff

    # ---------------- 3) Nurse satisfaction (as cost) ----------------
    # Include:
    # - Late followed by Early penalty
    # - More than 5 consecutive working days
    # - Wrong total shifts vs 20/15/10 rule
    # - Individual preferences pref[n][d][s] (only when nurse works at all)

    LATE_SHIFT = 2   # from your encoding: 0=E,1=D,2=L,3=N,4=F
    EARLY_SHIFT = 0
    LATE_EARLY_PEN = 50.0   # high penalty for L->E

    CONS_WORK_LIMIT = 5     # "no nurse should work more than 5 days in a row"
    CONS_WORK_PEN   = 50.0

    ASSIGN_PEN      = 10.0  # penalty per deviation from target shifts
    PREF_PEN        = 1.0   # scalar on preference matrix

    for n in range(number_nurses):
        works_anything = any(roster[n][d] < 4 for d in range(number_days))
        if not works_anything:
            # no wages, no penalties for this nurse
            continue

        # (1) Late → Early transitions
        for d in range(1, number_days):
            prev_s = roster[n][d - 1]
            curr_s = roster[n][d]
            if prev_s == LATE_SHIFT and curr_s == EARLY_SHIFT:
                nurse_cost += LATE_EARLY_PEN

        # (2) Max 5 consecutive working days
        cons = 0
        for d in range(number_days + 1):
            if d < number_days and roster[n][d] < 4:
                cons += 1
            else:
                if cons > CONS_WORK_LIMIT:
                    nurse_cost += CONS_WORK_PEN * (cons - CONS_WORK_LIMIT)
                cons = 0

        # (3) Target number of shifts based on employment
        # FTE -> 20, 0.75 -> 15, 0.5 -> 10
        emp = nurse_percent_employment[n]  # e.g. 1.0, 0.75, 0.5
        target_shifts = 20.0 * emp

        worked = sum(1 for d in range(number_days) if roster[n][d] < 4)
        nurse_cost += ASSIGN_PEN * abs(worked - target_shifts)

        # (4) Individual preferences (only when working)
        for d in range(number_days):
            s = roster[n][d]
            if s < 4:  # working shift
                nurse_cost += PREF_PEN * pref[n][d][s]

    return wage_cost, nurse_cost, patient_cost


def compute_objective(roster):
    """
    Weighted objective: 0.2 * totalWages + 10 * NurseSatisfaction + 2 * PatientSatisfaction
    All three components are already costs (higher = worse).
    """
    wage_cost, nurse_cost, patient_cost = compute_components(roster)
    return (
        WEIGHT_WAGE    * wage_cost +
        WEIGHT_NURSE   * nurse_cost +
        WEIGHT_PATIENT * patient_cost
    )

def random_neighbor(roster):
    """
    Make a small change:
    - pick a random day
    - pick 2 random nurses
    - swap their assignments on that day

    Returns a NEW roster (deep copy).
    """
    new_roster = deepcopy(roster)

    if number_nurses < 2:
        return new_roster

    d = random.randrange(number_days)
    n1 = random.randrange(number_nurses)
    n2 = random.randrange(number_nurses)
    while n2 == n1:
        n2 = random.randrange(number_nurses)

    s1 = new_roster[n1][d]
    s2 = new_roster[n2][d]
    new_roster[n1][d] = s2
    new_roster[n2][d] = s1

    return new_roster

def simulated_annealing(initial_roster,
                        T_start=1000.0,
                        T_min=1e-3,
                        alpha=0.95,
                        iters_per_T=200):
    """
    Standard simulated annealing:
    - start from initial_roster
    - occasionally accept worse moves with probability exp(-delta/T)
    - gradually cool down
    """
    current = deepcopy(initial_roster)
    best = deepcopy(initial_roster)

    current_cost = compute_objective(current)
    best_cost = current_cost

    T = T_start

    while T > T_min:
        for _ in range(iters_per_T):
            neighbor = random_neighbor(current)
            neighbor_cost = compute_objective(neighbor)
            delta = neighbor_cost - current_cost

            if delta < 0:
                # better → accept always
                current = neighbor
                current_cost = neighbor_cost
                if neighbor_cost < best_cost:
                    best = deepcopy(neighbor)
                    best_cost = neighbor_cost
            else:
                # worse → accept with probability
                p = math.exp(-delta / T)
                if random.random() < p:
                    current = neighbor
                    current_cost = neighbor_cost

        T *= alpha  # cool down

    return best, best_cost



def procedure():
    """
    Construct the monthly roster for department A.

    Steps:
    1) Read the input schedule from Excel (Case_E_MonthlyRoster_A).
    2) Use simulated annealing to improve it w.r.t. weighted objective:
         0.2 * Wage + 10 * NurseSat + 2 * PatientSat
    3) Store the best roster back into global monthly_roster.
    """
    global monthly_roster

    # 1) start from the Excel monthly roster
    read_monthly_roster_from_excel()

    # convert to plain list-of-lists for SA
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

    # 2) run simulated annealing
    best_roster, best_obj = simulated_annealing(initial_roster)

    w1, n1, p1 = compute_components(best_roster)
    print("Best schedule metrics after SA:")
    print(f"  Wage_cost      = {w1:.2f}")
    print(f"  Nurse_cost     = {n1:.2f}")
    print(f"  Patient_cost   = {p1:.2f}")
    print(f"  Objective      = {best_obj:.2f}")

    # 3) copy best_roster back into the global monthly_roster
    for n in range(number_nurses):
        for d in range(number_days):
            monthly_roster[n][d] = best_roster[n][d]


def add_nurse_to_day_shift(nurse_id: int, day_id: int, shift_id: int):
    """
    Assign nurse `nurse_id` to `shift_id` on `day_id` in the internal encoding.
    """
    monthly_roster[nurse_id][day_id] = shift_id


import time
import random


def main():
    global number_days, weekend, department, elapsed_time

    # GENERAL CHARACTERISTICS
    number_days = 28           # planning horizon
    weekend = 7                # first Sunday on day 7
    department = "A"           # adapt if needed: "A", "B", "C", "D"

    # INITIALISATION
    seed = 1000
    random.seed(seed)

    debug_list_sheets()

    # READ INPUT
    read_input()

    # Construct monthly roster and measure time
    start_time = time.perf_counter()
    procedure()
    elapsed_time = time.perf_counter() - start_time
    print(f"CPU time for procedure(): {elapsed_time:.6f} seconds")

    # 1) TXT + DataFrame for monthly roster
    df_roster = print_output()

    # 2) TXT + DataFrames for evaluation
    df_summary, df_staffing = evaluate_solution()

    # 3) Write everything into ONE Excel file with multiple sheets
    output_file = BASE_DIR / f"CASE_E_output_{department}.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_roster.to_excel(writer, sheet_name="MonthlyRoster", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_staffing.to_excel(writer, sheet_name="StaffingViolations", index=False)

    print(f"\nExcel output written to: {output_file}")

if __name__ == "__main__":
    main()


