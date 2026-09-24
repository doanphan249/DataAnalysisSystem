"""Kiểu dữ liệu cấu hình, kết quả và nhật ký dùng chung."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
import math


class DataError(ValueError):
    """Dữ liệu nguồn thiếu/sai; dừng thay vì tạo kết quả có vẻ hợp lệ."""


@dataclass(frozen=True)
class Settings:
    source_folder: Path
    output_file: Path
    year: int | None = None
    include_previous_december: bool = True
    apply_reference_overrides: bool = True


@dataclass(frozen=True)
class MonthFiles:
    year: int
    month: int
    production: Path
    sales: Path


@dataclass(frozen=True)
class RunResult:
    output_file: Path
    audit_file: Path
    business_rows: int
    warnings: int


@dataclass
class Report:
    schemas: dict[str, list[str]]
    rows: dict[str, list[list[Any]]] = field(default_factory=dict)
    lineage: dict[str, list[str]] = field(default_factory=dict)
    warnings: list[dict] = field(default_factory=list)
    checks: list[dict] = field(default_factory=list)

    def __post_init__(self):
        self.rows = {name: [] for name in self.schemas}
        self.lineage = {name: [] for name in self.schemas}

    def add(self, sheet: str, values: list, source: str) -> None:
        if len(values) != len(self.schemas[sheet]):
            raise DataError(f'Sai số cột {sheet}: {len(values)}')
        self.rows[sheet].append(values)
        self.lineage[sheet].append(source)

    def reorder(self, sheet: str, sort_key) -> None:
        """Sắp ổn định các dòng của sheet; lineage đi cùng dòng để vẫn khớp hàng Excel."""
        pairs = sorted(zip(self.rows[sheet], self.lineage[sheet]), key=lambda pair: sort_key(pair[0]))
        self.rows[sheet] = [row for row, _ in pairs]
        self.lineage[sheet] = [source for _, source in pairs]

    def warn(self, code: str, message: str, source: str='') -> None:
        item = {'code': code, 'message': message, 'source': source}
        if item not in self.warnings:
            self.warnings.append(item)

    def check(self, name: str, actual: float, expected: float, tolerance: float=1e-06, source: str='') -> None:
        difference = actual - expected
        ok = math.isfinite(difference) and abs(difference) <= tolerance
        self.checks.append(dict(name=name, actual=actual, expected=expected, difference=difference, tolerance=tolerance, passed=ok, source=source))
        if not ok:
            raise DataError(f'Đối chiếu không khớp: {name}; thực tế={actual}, nguồn={expected}; {source}')
