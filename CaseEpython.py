#Test Axel
import pip
import ortools

from ortools.sat.python import cp_model
print("ORTools ok")



#parameters

from dataclasses import dataclass
from pyexpat import model


@dataclass
class Nurse:
    id: str
    type: int
    employment_fraction: float
    fixed_leaves: list
    preferred_shifts: dict  # {('Monday', 'E'): True, ...}
    max_shifts_per_week: int = 6

@dataclass
class Required:
    day: str
    shift: str
    type: int
    requirement: int

days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
shifts = ["E", "D", "L", "N" ]
nurses: list[Nurse] = [...]
nurse_ids = [n.id for n in nurses]
nurse_types = sorted(set(n.type for n in nurses))

Max_Shifts_Per_Day = 1
Shift_Length = {'E': 9, 'D': 9 ,'L': 9, 'N': 12}
TOTAL_WEEKLY_HOURS = 38

requirement = {...}

model = cp_model.CpModel()

# x[n_id, d, s] = 1 if nurse n works shift s on day d
x = {}
for nurse in nurses:
    for d in days:
        for s in shifts:
            x[(nurse.id, d, s)] = model.NewBoolVar(f"x_{nurse.id}_{d}_{s}")

#Preference penalties

PENALTY = {
    "wanted": 0,
    "neutral": 1,
    "unwanted": 5,
    "forbidden": 1000,
}

def get_pref_label(nurse: Nurse, day: str, shift: str) -> str:
    # default if not specified
    return nurse.preferred_shifts.get((day, shift), "neutral")

def get_penalty(nurse: Nurse, day: str, shift: str) -> int:
    return PENALTY[get_pref_label(nurse, day, shift)]


#A nurse can work at most one shift per day
for nurse in nurses:
    for d in days:
        model.Add(
            sum(x[(nurse.id, d, s)] for s in shifts) <= Max_Shifts_Per_Day
        )



#Leave constraints
for nurse in nurses:
    for d in nurse.fixed_leaves:
        for s in shifts:
            model.Add(x[(nurse.id, d, s)] == 0)



#Weekly workload constraint
for nurse in nurses:
    # Hour budget
    model.Add(
        sum(
            x[(nurse.id, d, s)] * Shift_Length[s]
            for d in days
            for s in shifts
        ) <= int(nurse.employment_fraction * TOTAL_WEEKLY_HOURS)
    )

    # Max shifts/week
    model.Add(
        sum(
            x[(nurse.id, d, s)]
            for d in days
            for s in shifts
        ) <= nurse.max_shifts_per_week
    )



#Coverage constraints
for d in days:
    for s in shifts:
        for t in nurse_types:
            model.Add(
                sum(
                    x[(n.id, d, s)]
                    for n in nurses
                    if n.type == t
                ) >= demand[d][s][t]
            )


#Objective

penalty_terms = []
for nurse in nurses:
    for d in days:
        for s in shifts:
            coeff = get_penalty(nurse, d, s)
            if coeff != 0:
                penalty_terms.append(coeff * x[(nurse.id, d, s)])

model.Minimize(sum(penalty_terms))


solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 30  # adjust if needed

status = solver.Solve(model)

if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    print(f"Status: {solver.StatusName(status)}")
    for nurse in nurses:
        print(f"\nSchedule for {nurse.id}:")
        for d in days:
            assigned = [s for s in shifts if solver.Value(x[(nurse.id, d, s)]) == 1]
            if assigned:
                print(f"  {d}: {assigned[0]}")
else:
    print("No feasible solution found.")

