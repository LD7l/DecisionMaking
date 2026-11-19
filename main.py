"""Build and evaluate the example monthly nurse roster."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence


DAYS = 28
SHIFT_COUNT = 5
WORKING_SHIFTS = 4
NURSE_TYPES = 2
OFF_SHIFT = 4


class InputFormatError(ValueError):
    """Raised when an input file does not match the expected format."""


class TokenReader:
    def __init__(self, path: Path) -> None:
        self.path = path
        try:
            self.tokens = path.read_text(encoding="utf-8").split()
        except OSError as error:
            raise OSError(f"Cannot read input file '{path}': {error}") from error
        self.position = 0

    def _next(self, description: str) -> str:
        if self.position >= len(self.tokens):
            raise InputFormatError(
                f"Input file '{self.path}' ended while reading {description}."
            )
        token = self.tokens[self.position]
        self.position += 1
        return token

    def read_int(self, description: str) -> int:
        token = self._next(description)
        try:
            return int(token)
        except ValueError as error:
            raise InputFormatError(
                f"Expected an integer for {description} in '{self.path}', got {token!r}."
            ) from error

    def read_float(self, description: str) -> float:
        token = self._next(description)
        try:
            return float(token)
        except ValueError as error:
            raise InputFormatError(
                f"Expected a number for {description} in '{self.path}', got {token!r}."
            ) from error

    def read_text(self, description: str) -> str:
        return self._next(description)


@dataclass
class ShiftSystem:
    raw_to_shift: list[int]
    requirements: list[int]


@dataclass
class Nurse:
    personnel_number: str
    preferences: list[list[int]]
    employment: float
    nurse_type: int
    cyclic_roster: list[int]
    min_assignments: int = 0
    max_assignments: int = 0
    min_consecutive_work: int = 0
    max_consecutive_work: int = 0
    min_consecutive_shift: list[int] = field(
        default_factory=lambda: [0] * SHIFT_COUNT
    )
    max_consecutive_shift: list[int] = field(
        default_factory=lambda: [0] * SHIFT_COUNT
    )
    min_shift_assignments: list[int] = field(default_factory=lambda: [0] * SHIFT_COUNT)
    max_shift_assignments: list[int] = field(default_factory=lambda: [0] * SHIFT_COUNT)
    identical_weekends: bool = False


def input_path(data_dir: Path, filename: str, department: str) -> Path:
    return data_dir / f"{filename}_dpt_{department}.txt"


def read_shift_system(data_dir: Path, department: str) -> ShiftSystem:
    path = input_path(data_dir, "Shift_system", department)
    reader = TokenReader(path)
    number_of_shifts = reader.read_int("number of shifts")
    reader.read_int("shift length")
    if not 1 <= number_of_shifts < SHIFT_COUNT:
        raise InputFormatError(
            f"'{path}' must describe between 1 and {SHIFT_COUNT - 1} working shifts."
        )

    start_times = [
        reader.read_int(f"start time for shift {shift_id}")
        for shift_id in range(1, number_of_shifts + 1)
    ]
    raw_to_shift = [OFF_SHIFT] + [OFF_SHIFT] * number_of_shifts
    requirements = [0] * SHIFT_COUNT

    for shift_id, start_time in enumerate(start_times, start=1):
        requirement = reader.read_int(f"staffing requirement for shift {shift_id}")
        if 3 <= start_time < 9:
            target = 0 if requirements[0] == 0 else 1
        elif 9 <= start_time < 12:
            target = 1 if requirements[1] == 0 else 2
        elif 12 <= start_time < 21:
            target = 2 if requirements[2] == 0 else 3
        elif start_time >= 21 or start_time < 3:
            target = 3
            if requirements[target] != 0:
                raise InputFormatError(
                    f"'{path}' contains more than one night shift."
                )
        else:
            raise InputFormatError(
                f"Shift {shift_id} in '{path}' has unsupported start time "
                f"{start_time}."
            )

        raw_to_shift[shift_id] = target
        requirements[target] = requirement

    return ShiftSystem(raw_to_shift, requirements)


def read_nurses(
    data_dir: Path, department: str, shift_system: ShiftSystem
) -> list[Nurse]:
    cyclic_path = input_path(data_dir, "Cyclic_roster", department)
    cyclic_reader = TokenReader(cyclic_path)
    number_of_nurses = cyclic_reader.read_int("number of nurses")
    if number_of_nurses <= 0:
        raise InputFormatError(f"'{cyclic_path}' must contain at least one nurse.")

    cyclic_rosters: list[list[int]] = []
    for nurse_index in range(number_of_nurses):
        roster = []
        for day in range(DAYS):
            raw_shift = cyclic_reader.read_int(
                f"shift for nurse {nurse_index + 1}, day {day + 1}"
            )
            if not 0 <= raw_shift < len(shift_system.raw_to_shift):
                raise InputFormatError(
                    f"'{cyclic_path}' contains invalid shift code {raw_shift}."
                )
            roster.append(shift_system.raw_to_shift[raw_shift])
        cyclic_rosters.append(roster)

    personnel_path = input_path(data_dir, "Personnel", department)
    personnel_reader = TokenReader(personnel_path)
    nurses = []
    for nurse_index, cyclic_roster in enumerate(cyclic_rosters):
        personnel_number = personnel_reader.read_text(
            f"personnel number for nurse {nurse_index + 1}"
        )
        preferences = [
            [
                personnel_reader.read_int(
                    f"preference for nurse {nurse_index + 1}, "
                    f"day {day + 1}, shift {shift}"
                )
                for shift in range(SHIFT_COUNT)
            ]
            for day in range(DAYS)
        ]
        employment = personnel_reader.read_float(
            f"employment percentage for nurse {nurse_index + 1}"
        )
        nurse_type = personnel_reader.read_int(
            f"nurse type for nurse {nurse_index + 1}"
        ) - 1
        if not 0 <= nurse_type < NURSE_TYPES:
            raise InputFormatError(
                f"'{personnel_path}' gives nurse {nurse_index + 1} "
                f"an unsupported nurse type."
            )
        nurses.append(
            Nurse(
                personnel_number=personnel_number,
                preferences=preferences,
                employment=employment,
                nurse_type=nurse_type,
                cyclic_roster=cyclic_roster,
            )
        )

    return nurses


def read_roster_rules(
    data_dir: Path, department: str, nurses: list[Nurse], shift_system: ShiftSystem
) -> None:
    path = input_path(data_dir, "Constraints", department)
    reader = TokenReader(path)
    min_assignments = reader.read_int("minimum number of assignments")
    max_assignments = reader.read_int("maximum number of assignments")
    min_consecutive_work = reader.read_int("minimum consecutive working days")
    max_consecutive_work = reader.read_int("maximum consecutive working days")

    min_consecutive_shift = [0] * SHIFT_COUNT
    max_consecutive_shift = [0] * SHIFT_COUNT
    min_shift_assignments = [0] * SHIFT_COUNT
    max_shift_assignments = [0] * SHIFT_COUNT
    for raw_shift in range(1, len(shift_system.raw_to_shift)):
        internal_shift = shift_system.raw_to_shift[raw_shift]
        min_consecutive_shift[internal_shift] = reader.read_int(
            f"minimum consecutive assignments for shift {raw_shift}"
        )
        max_consecutive_shift[internal_shift] = reader.read_int(
            f"maximum consecutive assignments for shift {raw_shift}"
        )

    for raw_shift in range(1, len(shift_system.raw_to_shift)):
        internal_shift = shift_system.raw_to_shift[raw_shift]
        min_shift_assignments[internal_shift] = reader.read_int(
            f"minimum assignments for shift {raw_shift}"
        )
        max_shift_assignments[internal_shift] = reader.read_int(
            f"maximum assignments for shift {raw_shift}"
        )

    identical_weekends = reader.read_text("identical weekend rule").upper().startswith(
        "Y"
    )
    for nurse in nurses:
        nurse.min_assignments = int(min_assignments * nurse.employment)
        nurse.max_assignments = int(max_assignments * nurse.employment)
        nurse.min_consecutive_work = min_consecutive_work
        nurse.max_consecutive_work = max_consecutive_work
        nurse.min_consecutive_shift = min_consecutive_shift.copy()
        nurse.max_consecutive_shift = max_consecutive_shift.copy()
        nurse.min_shift_assignments = min_shift_assignments.copy()
        nurse.max_shift_assignments = max_shift_assignments.copy()
        nurse.identical_weekends = identical_weekends


def build_roster(nurses: list[Nurse]) -> list[list[int]]:
    return [nurse.cyclic_roster.copy() for nurse in nurses]


def write_roster(
    data_dir: Path,
    department: str,
    nurses: list[Nurse],
    roster: list[list[int]],
    shift_system: ShiftSystem,
) -> None:
    output_path = input_path(data_dir, "Monthly_Roster", department)
    with output_path.open("w", encoding="utf-8", newline="\n") as output:
        for nurse, assignments in zip(nurses, roster):
            output.write(f"{nurse.personnel_number}\t")
            for assignment in assignments:
                raw_shift = shift_system.raw_to_shift.index(assignment)
                output.write(f"{raw_shift}\t")
            output.write("\n")


def evaluate_solution(
    data_dir: Path,
    department: str,
    nurses: list[Nurse],
    roster: list[list[int]],
    shift_system: ShiftSystem,
) -> list[str]:
    scheduled = [[0] * SHIFT_COUNT for _ in range(DAYS)]
    preference_score = 0
    max_work_violations = 0
    max_shift_violations = 0
    min_assignment_violations = 0
    max_assignment_violations = 0

    for nurse, assignments in zip(nurses, roster):
        work_run = 0
        shift_run = 0
        previous_shift = OFF_SHIFT
        total_assignments = 0

        for day, shift in enumerate(assignments):
            if not 0 <= shift < SHIFT_COUNT:
                raise InputFormatError(
                    f"Nurse {nurse.personnel_number} has invalid shift code {shift}."
                )
            preference_score += nurse.preferences[day][shift]
            scheduled[day][shift] += 1
            if shift < WORKING_SHIFTS:
                total_assignments += 1
                work_run += 1
                if shift == previous_shift:
                    shift_run += 1
                else:
                    if (
                        previous_shift < WORKING_SHIFTS
                        and shift_run > nurse.max_consecutive_shift[previous_shift]
                    ):
                        max_shift_violations += 1
                    shift_run = 1
            else:
                if work_run > nurse.max_consecutive_work:
                    max_work_violations += 1
                if (
                    previous_shift < WORKING_SHIFTS
                    and shift_run > nurse.max_consecutive_shift[previous_shift]
                ):
                    max_shift_violations += 1
                work_run = 0
                shift_run = 0
            previous_shift = shift

        if work_run > nurse.max_consecutive_work:
            max_work_violations += 1
        if (
            previous_shift < WORKING_SHIFTS
            and shift_run > nurse.max_consecutive_shift[previous_shift]
        ):
            max_shift_violations += 1
        if total_assignments < nurse.min_assignments:
            min_assignment_violations += 1
        if total_assignments > nurse.max_assignments:
            max_assignment_violations += 1

    lines = [
        f"The total preference score is {preference_score}.",
        "The constraint 'maximum number of consecutive working days' "
        f"is violated {max_work_violations} times.",
        "The constraint 'maximum number of consecutive working days per shift type' "
        f"is violated {max_shift_violations} times.",
        "The constraint 'minimum number of assignments' "
        f"is violated {min_assignment_violations} times.",
        "The constraint 'maximum number of assignments' "
        f"is violated {max_assignment_violations} times.",
        "",
        "The staffing requirements are violated as follows:",
    ]
    for day in range(DAYS):
        for shift in range(WORKING_SHIFTS):
            assigned = scheduled[day][shift]
            raw_shift = next(
                (
                    raw_code
                    for raw_code, internal_code in enumerate(
                        shift_system.raw_to_shift
                    )
                    if internal_code == shift
                ),
                shift,
            )
            if assigned < shift_system.requirements[shift]:
                lines.append(
                    f"There are too few nurses in shift {raw_shift} on day {day}: "
                    f"{assigned} < {shift_system.requirements[shift]}."
                )
            elif assigned > shift_system.requirements[shift]:
                lines.append(
                    f"There are too many nurses in shift {raw_shift} on day {day}: "
                    f"{assigned} > {shift_system.requirements[shift]}."
                )

    output_path = input_path(data_dir, "Violations", department)
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return lines


def run(data_dir: Path, department: str) -> None:
    shift_system = read_shift_system(data_dir, department)
    nurses = read_nurses(data_dir, department, shift_system)
    read_roster_rules(data_dir, department, nurses, shift_system)
    roster = build_roster(nurses)
    write_roster(data_dir, department, nurses, roster, shift_system)
    evaluate_solution(data_dir, department, nurses, roster, shift_system)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build and evaluate the cyclic nurse roster."
    )
    parser.add_argument(
        "--department",
        choices=("A", "B", "C", "D"),
        default="A",
        help="department input files to use (default: A)",
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=Path("files"),
        help="directory containing the department input files (default: files)",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        run(args.data_dir, args.department)
    except (OSError, InputFormatError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
