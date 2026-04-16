from dataclasses import dataclass

@dataclass
class Student:
    student_id: str
    student_name: str
    grade: float | None = None


@dataclass
class HistoryEntry:
    student_id: str
    student_name: str
    grade: float

