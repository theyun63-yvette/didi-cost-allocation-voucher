from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime
from decimal import Decimal
from typing import Any


@dataclass
class Issue:
    code: str
    level: str
    blocking: bool
    file_name: str = ""
    sheet: str = ""
    row: int | None = None
    order_id: str = ""
    employee_name: str = ""
    position: str = ""
    original_value: Any = ""
    reason: str = ""
    suggestion: str = ""

    def as_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["blocking"] = "是" if self.blocking else "否"
        return data


@dataclass
class PersonnelRecord:
    name: str
    position: str
    employee_id: str = ""
    department_code: str = ""
    department_name: str = ""
    source_sheet: str = ""
    source_row: int = 0


@dataclass
class AllocationRecord:
    name: str
    position: str
    project_ratios: dict[str, Decimal]
    non_rnd_ratio: Decimal
    source_sheet: str
    source_row: int
    valid: bool = True
    non_rnd_missing: bool = False
    blank_projects: tuple[str, ...] = ()


@dataclass
class DidiOrder:
    order_id: str
    employee_name: str
    enterprise_paid: Decimal
    refund_amount: Decimal
    order_datetime: datetime | None
    source_sheet: str
    source_row: int
    status: str = ""
    raw: dict[str, Any] = field(default_factory=dict)
    valid: bool = True
    raw_employee_name: str = ""
    billing_period: str = ""
    billing_month: str = ""
    original_enterprise_paid: Decimal | None = None
    enterprise_payable: Decimal | None = None
    refund_datetime: datetime | None = None
    refund_handling: str = "无退款"


@dataclass
class AllocationDetail:
    order_id: str
    employee_name: str
    position: str
    gross_amount: Decimal
    order_net_amount: Decimal
    allocation_type: str
    project_name: str
    project_ratio: Decimal
    allocation_amount: Decimal
    account_code: str
    account_name: str
    department_code: str
    department_name: str
    project_code: str
    currency_code: str
    source_sheet: str
    source_row: int

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class VoucherLine:
    line_type: str
    account_code: str
    account_name: str
    debit: Decimal = Decimal("0")
    credit: Decimal = Decimal("0")
    department_code: str = ""
    department_name: str = ""
    project_code: str = ""
    project_name: str = ""
    dimension_type: str = ""
    dimension_code: str = ""
    dimension_name: str = ""
    tail_adjustment: bool = False

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ParseResult:
    records: list[Any] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class ProcessingResult:
    target_month: str
    issues: list[Issue]
    details: list[AllocationDetail]
    match_rows: list[dict[str, Any]]
    summary_rows: list[dict[str, Any]]
    voucher_lines: list[VoucherLine]
    metrics: dict[str, Any]
    recognized_sheets: dict[str, str]
    audit_log: list[dict[str, Any]] = field(default_factory=list)
    output_files: dict[str, bytes] = field(default_factory=dict)
    order_detail_rows: list[dict[str, Any]] = field(default_factory=list)
    cleaning_logs: list[dict[str, Any]] = field(default_factory=list)
    valid_orders: list[DidiOrder] = field(default_factory=list)

    @property
    def blocking_issues(self) -> list[Issue]:
        return [x for x in self.issues if x.blocking]

    @property
    def can_generate_voucher(self) -> bool:
        return not self.blocking_issues
