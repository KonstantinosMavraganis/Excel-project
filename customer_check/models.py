from dataclasses import dataclass, field


@dataclass(frozen=True)
class Row:
    number: int
    values: tuple[str, ...]


@dataclass(frozen=True)
class Table:
    source: str
    sheet: str
    headers: tuple[str, ...]
    rows: tuple[Row, ...]


@dataclass(frozen=True)
class Entry:
    code: str
    name: str
    source_row: int
    detail: str = ""


@dataclass
class Result:
    found: list[Entry] = field(default_factory=list)
    missing: list[Entry] = field(default_factory=list)
    unknown: list[Entry] = field(default_factory=list)
    issues: list[Entry] = field(default_factory=list)
    registry_rows: int = 0
    payment_rows: int = 0
    inactive_rows: int = 0
    duplicate_payments: int = 0

