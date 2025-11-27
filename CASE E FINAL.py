import math
import time
import random
import pandas as pd
from pathlib import Path

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


def read_shift_system():
    """
    Read the shift system for the current `department` into the global structures.

    Expects a file: files/Shift_system_dpt_<department>.txt
    Format (same as C++ version expects):
      - first line: <number_shifts>\t<length>
      - next `number_shifts` lines: start time (int) of each shift
      - then: required staff numbers, as integers separated by whitespace
    """
    global number_shifts, length, hrs, req, shift, start_shift, end_shift

    # Build filename, same as C++: "files/Shift_system_dpt_" + department + ".txt"
    filename = f"Shift_system_dpt_{department}.txt"
    


    with open(filename, "r") as f:
        tokens = f.read().split()   # split on ANY whitespace

    p = 0  # pointer in tokens

    # number_shifts and length
    number_shifts = int(tokens[p]); p += 1
    length = int(tokens[p]); p += 1

    # start times for the different shifts (1-based indexing like C++)
    for k in range(1, number_shifts + 1):
        start_shift[k] = int(tokens[p]); p += 1

    # the remaining ints are the requirements (what C++ read with those fscanf("%d\t",...) calls)
    req_values = list(map(int, tokens[p:]))
    q = 0  # pointer in req_values

    i = 0  # day 0 pattern

    # SHIFT ENCODING:
    # Early  -> code 0
    # Day    -> code 1
    # Late   -> code 2
    # Night  -> code 3
    # Day off-> code 4 (assigned to shift[0])
    for k in range(1, number_shifts + 1):
        start_k = start_shift[k]

        # EARLY (3 <= start < 9)
        if 3 <= start_k < 9 and req[i][0] == 0:
            req[i][0] = req_values[q]; q += 1
            shift[k] = 0
        elif 3 <= start_k < 9 and req[i][0] != 0:
            req[i][1] = req_values[q]; q += 1
            shift[k] = 1

        # DAY (9 <= start < 12)
        if 9 <= start_k < 12 and req[i][1] == 0:
            req[i][1] = req_values[q]; q += 1
            shift[k] = 1
        elif 9 <= start_k < 12 and req[i][1] != 0:
            req[i][2] = req_values[q]; q += 1
            shift[k] = 2

        # LATE (12 <= start < 21)
        if 12 <= start_k < 21 and req[i][2] == 0:
            req[i][2] = req_values[q]; q += 1
            shift[k] = 2
        elif 12 <= start_k < 21 and req[i][2] != 0:
            req[i][3] = req_values[q]; q += 1
            shift[k] = 3

        # NIGHT (start >= 21 or start < 3)
        if ((start_k >= 21 or start_k < 3) and req[i][3] == 0):
            req[i][3] = req_values[q]; q += 1
            shift[k] = 3
        elif ((start_k >= 21 or start_k < 3) and req[i][3] != 0):
            print("Read problem shifts input")

    # Day off associated with shift 4
    shift[0] = 4

    # Determine end times and hrs for each shift
    for m in range(1, number_shifts + 1):
        if start_shift[m] + length < 24:
            hrs[shift[m]] = length
            end_shift[m] = start_shift[m] + length
        else:
            hrs[shift[m]] = length
            end_shift[m] = hrs[shift[m]] + start_shift[m] - 24

    # Free shift (day off) has zero hours
    hrs[shift[0]] = 0

    # Copy staffing requirements to the other days
    for day in range(1, number_days):
        for j in range(0, number_shifts + 1):  # j=0..number_shifts, using shift[j]
            req[day][shift[j]] = req[0][shift[j]]

    # Include day off as an extra shift
    number_shifts += 1

def read_personnel_characteristics():
    """
    Read preferences and characteristics for all nurses for this department.
    Mirrors the C++ read_personnel_characteristics().
    """
    global number_types, personnel_number, pref, nurse_percent_employment, nurse_type

    filename = f"Personnel_dpt_{department}.txt"
    number_types = TYPES  # same constant as C++

    with open(filename, "r") as f:
        tokens = f.read().split()

    p = 0  # pointer into tokens

    for k in range(number_nurses):
        # Personnel number (string)
        personnel_number[k] = tokens[p]
        p += 1

        # Preferences: number_days * 5
        for i in range(number_days):
            for j in range(5):  # file always contains 5 shift-pref columns
                pref[k][i][j] = int(tokens[p])
                p += 1

        # Percentage of employment and nurse type
        nurse_percent_employment[k] = float(tokens[p]); p += 1
        nurse_type[k] = int(tokens[p]) - 1; p += 1

def read_cyclic_roster():
    """
    Read the cyclic roster for this department.
    Debug version: checks token counts and prints diagnostics.
    """
    global number_nurses, cyclic_roster

    filename = f"Cyclic_roster_dpt_{department}.txt"

    with open(filename, "r") as f:
        content = f.read()
    tokens = content.split()

    print("---- DEBUG read_cyclic_roster ----")
    print("Raw token count:", len(tokens))

    if not tokens:
        raise ValueError("Cyclic roster file is empty")

    p = 0
    number_nurses = int(tokens[p]); p += 1
    print("number_nurses from file:", number_nurses)
    print("number_days in code:", number_days)

    expected_len = 1 + number_nurses * number_days
    print("expected total tokens:", expected_len)

    if len(tokens) < expected_len:
        raise ValueError(
            f"Not enough data in cyclic roster file: "
            f"have {len(tokens)} tokens, expected at least {expected_len}"
        )

    # if counts are fine, fill the roster
    for k in range(number_nurses):
        for i in range(number_days):
            l = int(tokens[p]); p += 1
            cyclic_roster[k][i] = shift[l]

    print("Cyclic roster read OK.")
    print("-------------------------------")

def read_monthly_roster_rules():
    """
    Read monthly roster rules for this department.
    Mirrors the C++ read_monthly_roster_rules().
    """
    global min_ass, max_ass, min_cons_wrk, max_cons_wrk
    global min_cons, max_cons, extreme_max_cons, extreme_min_cons
    global min_shift, max_shift, identical
    global extreme_max_cons_wrk, extreme_min_cons_wrk

    filename = f"Constraints_dpt_{department}.txt"

    for k in range(number_nurses):
        with open(filename, "r") as f:
            tokens = f.read().split()

        p = 0

        # min/max assignments over whole period
        min_ass[k] = int(tokens[p]); p += 1
        max_ass[k] = int(tokens[p]); p += 1

        # scale by employment fraction (mimic C++ int *= float truncation)
        min_ass[k] = int(min_ass[k] * nurse_percent_employment[k])
        max_ass[k] = int(max_ass[k] * nurse_percent_employment[k])

        # min/max consecutive work days
        min_cons_wrk[k] = int(tokens[p]); p += 1
        max_cons_wrk[k] = int(tokens[p]); p += 1

        extreme_max_cons_wrk = 10
        extreme_min_cons_wrk = 1

        # per-shift-type consecutive work day constraints
        for i in range(1, number_shifts):
            h1 = int(tokens[p]); p += 1
            h2 = int(tokens[p]); p += 1
            sh = shift[i]
            min_cons[k][sh] = h1
            max_cons[k][sh] = h2
            extreme_max_cons[k][sh] = 10
            extreme_min_cons[k][sh] = 1

        # min/max assignments per shift type over the whole period
        for i in range(1, number_shifts):
            sh = shift[i]
            min_shift[k][sh] = int(tokens[p]); p += 1
            max_shift[k][sh] = int(tokens[p]); p += 1

        # identical weekend constraint
        identical_string = tokens[p]  # e.g. "Y" or "N"
        if identical_string[0].upper() == "Y":
            identical[k] = 1
        else:
            identical[k] = 0

def read_input():
    """
    Read all input files and initialise data structures.
    Equivalent of C++ read_input().
    """
    global cyclic_roster, number_shifts

    # 1) shift system
    read_shift_system()

    # 2) initialise cyclic roster to 0
    for k in range(number_nurses):
        for i in range(number_days):
            cyclic_roster[k][i] = 0

    # 3) read the other input files
    read_cyclic_roster()
    read_personnel_characteristics()
    read_monthly_roster_rules()

    # 4) force number of shifts in algorithm to 5 (E, D, L, N, off)
    number_shifts = 5


def print_output():
    """
    Print the monthly roster in the student's original shift numbering.
    Equivalent of C++ print_output().
    """
    filename = f"Monthly_Roster_dpt_{department}.txt"

    with open(filename, "w") as f:
        for k in range(number_nurses):
            # personnel number
            f.write(f"{personnel_number[k]}\t")
            # daily shifts
            for i in range(number_days):
                shift_index = shift_decoding(monthly_roster[k][i])
                f.write(f"{shift_index}\t")
            f.write("\n")

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
    Evaluate the full personnel schedule and write violation report.
    Equivalent of C++ evaluate_solution().
    """
    global scheduled, violations

    filename = f"Violations_dpt_{department}.txt"

    # reset scheduled counts
    for kk in range(number_types):
        for i in range(number_days):
            for j in range(number_shifts):
                scheduled[kk][i][j] = 0

    # reset violations (first 20 entries)
    for i in range(20):
        violations[i] = 0

    # evaluate each nurse's line of work
    for i in range(number_nurses):
        evaluate_line_of_work(i, slack_j=0)

    with open(filename, "w") as f:
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
        for i in range(number_days):
            # last shift (day off) is usually excluded: 0..number_shifts-2
            for j in range(number_shifts - 1):
                a = 0
                for kk in range(number_types):
                    a += scheduled[kk][i][j]

                shift_index = shift_decoding(j)  # original shift index used in input

                if a < req[i][j]:
                    f.write(
                        f"There are too few nurses in shift {shift_index} on day {i}: "
                        f"{a} < {req[i][j]}.\n"
                    )
                elif a > req[i][j]:
                    f.write(
                        f"There are too many nurses in shift {shift_index} on day {i}: "
                        f"{a} > {req[i][j]}.\n"
                    )

def procedure():
    """
    Construct the monthly roster.
    Currently: dummy example, copy cyclic_roster into monthly_roster.
    """
    for k in range(number_nurses):
        for i in range(number_days):
            monthly_roster[k][i] = cyclic_roster[k][i]

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

    # READ INPUT
    read_input()

    # Construct monthly roster and measure time
    start_time = time.perf_counter()
    procedure()
    elapsed_time = time.perf_counter() - start_time
    print(f"CPU time for procedure(): {elapsed_time:.6f} seconds")

    # OUTPUT monthly roster
    print_output()

    # EVALUATE solution
    evaluate_solution()
    print("Evaluation written to Violations_dpt_*.txt")


if __name__ == "__main__":
    main()


