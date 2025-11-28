import math
import time
import random
import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

SHIFT_LABELS = {0: "E", 1: "D", 2: "L", 3: "N", 4: "F"}

# === CONSTANTS (replace the #define stuff) ===
NURSES = 100
DAYS = 30
SHIFTS = 5
TYPES = 2

# === GENERIC PERSONNEL ROSTERING VARIABLES ===
department: str = ""          # was: char department[10];
number_days: int = 0          # was: int number_days;
number_nurses: int = 0        # was: int number_nurses;
number_shifts: int = 0        # was: int number_shifts;
shift_code: int = 0           # was: int shift_code;


# === VARIABLES SHIFT SYSTEM ===
# int hrs[SHIFTS];
hrs = [0 for _ in range(SHIFTS)]

# int req[DAYS][SHIFTS];
req = [[0 for _ in range(SHIFTS)] for _ in range(DAYS)]

# int shift[SHIFTS];
shift = [0 for _ in range(SHIFTS)]

# int start_shift[SHIFTS];
start_shift = [0 for _ in range(SHIFTS)]

# int end_shift[SHIFTS];
end_shift = [0 for _ in range(SHIFTS)]

length: int = 0


# === VARIABLES PERSONNEL CHARACTERISTICS ===
number_types: int = 0                 # was: int number_types;

# int nurse_type[NURSES];
nurse_type = [0 for _ in range(NURSES)]

# int pref[NURSES][DAYS][SHIFTS];
pref = [
    [
        [0 for _ in range(SHIFTS)]
        for _ in range(DAYS)
    ]
    for _ in range(NURSES)
]

# float nurse_percent_employment[NURSES];
nurse_percent_employment = [0.0 for _ in range(NURSES)]

# std::string personnel_number[NURSES];
personnel_number = ["" for _ in range(NURSES)]


# === VARIABLES PERSONNEL ROSTER ===
# int cyclic_roster[NURSES][DAYS];
cyclic_roster = [
    [0 for _ in range(DAYS)]
    for _ in range(NURSES)
]

# int monthly_roster[NURSES][DAYS];
monthly_roster = [
    [0 for _ in range(DAYS)]
    for _ in range(NURSES)
]


# === VARIABLES MONTHLY ROSTER RULES ===
# int min_ass[NURSES];
min_ass = [0 for _ in range(NURSES)]

# int max_ass[NURSES];
max_ass = [0 for _ in range(NURSES)]

weekend: int = 0  # day the weekend starts

# int identical[NURSES];
identical = [0 for _ in range(NURSES)]

# int max_cons[NURSES][SHIFTS];
max_cons = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

# int min_cons[NURSES][SHIFTS];
min_cons = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

# int min_shift[NURSES][SHIFTS];
min_shift = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

# int max_shift[NURSES][SHIFTS];
max_shift = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

# int min_cons_wrk[NURSES];
min_cons_wrk = [0 for _ in range(NURSES)]

# int max_cons_wrk[NURSES];
max_cons_wrk = [0 for _ in range(NURSES)]

# int extreme_max_cons[NURSES][SHIFTS];
extreme_max_cons = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

# int extreme_min_cons[NURSES][SHIFTS];
extreme_min_cons = [
    [0 for _ in range(SHIFTS)]
    for _ in range(NURSES)
]

extreme_max_cons_wrk: int = 0
extreme_min_cons_wrk: int = 0


# === EVALUATION VARIABLES ===
count_ass: int = 0
count_cons_wrk: int = 0
count_cons: int = 0

# int count_shift[SHIFTS];
count_shift = [0 for _ in range(SHIFTS)]

# int scheduled[TYPES][DAYS][SHIFTS];
scheduled = [
    [
        [0 for _ in range(SHIFTS)]
        for _ in range(DAYS)
    ]
    for _ in range(TYPES)
]

# int violations[DAYS * SHIFTS];
violations = [0 for _ in range(DAYS * SHIFTS)]

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

def _find_cell_containing(df, text):
    """
    Return (row, col) of the first cell whose string contains 'text'
    (case-insensitive).
    """
    target = text.upper()
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = df.iat[r, c]
            if isinstance(val, str) and target in val.strip().upper():
                return r, c
    raise ValueError(f"Cell containing '{text}' not found in shift system sheet")


def read_shift_system():
    """
    Read the shift system for department A from the 'Case_E_Manual' sheet.

    Expected layout (simplified):

        A1: 'Number of shifts'     B1: 'Duration (in h)'
        A2: <num_shifts>          B2: <length>

        row ~4: 'Start shifts Dep A' | 'Start shifts Dep B' | ...

        below that (num_shifts rows): start hours for each shift, per dep

        row ~11: 'Requirements Dep A' | 'Requirements Dep B' | ...

        below that (num_shifts rows): required nurses per shift, per dep

    Internal encoding:
        0 = Early (E)   3 <= start < 9
        1 = Day   (D)   9 <= start < 12
        2 = Late  (L)   12 <= start < 21
        3 = Night (N)   start >= 21 or start < 3
        4 = Free  (F)
    """
    global number_shifts, length, hrs, req, shift, start_shift, end_shift, number_days

    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    sheet_name = "Case_C_9"   # adjust if your sheet has a different name

    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # ---- Number of shifts and length ----
    number_shifts = int(df.iat[1, 0])  # A2
    length = int(df.iat[1, 1])         # B2

    # ---- Find 'Start shifts Dep A' and 'Requirements Dep A' headers ----
    r_start, c_start = _find_cell_containing(df, "START SHIFTS DEP A")
    r_req, c_req = _find_cell_containing(df, "REQUIREMENTS DEP A")

    # The next 'number_shifts' rows after these headers contain numeric data
    start_rows = [r_start + 1 + i for i in range(number_shifts)]
    req_rows = [r_req + 1 + i for i in range(number_shifts)]

    # Reset requirements for day 0
    for j in range(SHIFTS):  # 0..4
        req[0][j] = 0

    # ---- Process each real shift (E/D/L) ----
    for idx in range(number_shifts):
        row_start = start_rows[idx]
        row_req = req_rows[idx]

        start_h = int(df.iat[row_start, c_start])
        required = int(df.iat[row_req, c_req])

        # 1-based index for compatibility with old arrays
        k = idx + 1
        start_shift[k] = start_h

        # decide internal code from start hour
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

        # compute end time
        if start_h + length < 24:
            end_shift[k] = start_h + length
        else:
            end_shift[k] = start_h + length - 24

    # ---- Free shift (code 4) ----
    shift[0] = 4
    hrs[4] = 0  # F has no hours

    # ---- Copy requirements to all days ----
    for day in range(1, number_days):
        for j in range(SHIFTS):
            req[day][j] = req[0][j]

    # We now have 5 shifts in the algorithm: E, D, L, N, F
    number_shifts = 5


import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent  # should already exist


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

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
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


def debug_list_sheets():
    excel_file = BASE_DIR / "CASE_E_input.xlsx"
    xls = pd.ExcelFile(excel_file)
    print("Excel file:", excel_file)
    print("Sheets:", xls.sheet_names)

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


def procedure():
    """
    Construct the monthly roster.

    Current implementation: read the schedule directly from Excel.
    """
    read_monthly_roster_from_excel()


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


