
#parameters

from dataclasses import dataclass

@dataclass
class Nurse:
    id: str
    type: int
    employment_fraction: float
    fixed_leaves: list
    preferred_shifts: dict  # {('Monday', 'E'): True, ...}
    max_shifts_per_week: int = 6