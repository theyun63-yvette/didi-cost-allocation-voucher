from __future__ import annotations

from typing import Any

from models.schemas import Issue, ParseResult
from .common import load_xlsx, normalize_text, source_name


def parse_voucher_template(source: Any) -> ParseResult:
    filename = source_name(source, "凭证标准模板.xlsx")
    try:
        wb = load_xlsx(source, data_only=False)
    except Exception as exc:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason=f"凭证模板无法打开: {exc}", suggestion="上传有效的xlsx模板")])
    candidates = []
    for ws in wb.worksheets:
        for row in range(1, min(ws.max_row, 15) + 1):
            vals = [normalize_text(ws.cell(row, c).value) for c in range(1, ws.max_column + 1)]
            if "book.number" in vals and "account.number" in vals and "debitlocal" in vals and "creditlocal" in vals:
                candidates.append((ws.title, row))
    if len(candidates) != 1:
        return ParseResult(issues=[Issue("E002", "错误", True, filename, reason="无法唯一识别凭证模板技术表头", suggestion="模板需包含 book.number、account.number、debitlocal、creditlocal")], metadata={"sheet_names": wb.sheetnames})
    sheet, technical_header_row = candidates[0]
    chinese_header_row = technical_header_row + 1
    data_start_row = chinese_header_row + 1
    ws = wb[sheet]
    tech_headers = [normalize_text(ws.cell(technical_header_row, c).value) for c in range(1, ws.max_column + 1)]
    required = ["book.number", "bizdate", "bookeddate", "vouchertype.number", "billno", "edescription", "account.number", "debitlocal", "creditlocal", "currency.number"]
    missing = [x for x in required if x not in tech_headers]
    issues = []
    if missing:
        issues.append(Issue("E002", "错误", True, filename, sheet, technical_header_row, original_value=", ".join(missing), reason="凭证模板缺少必要字段", suggestion="使用财务系统标准模板"))
    return ParseResult([], issues, {"sheet": sheet, "technical_header_row": technical_header_row, "chinese_header_row": chinese_header_row, "data_start_row": data_start_row, "sheet_names": wb.sheetnames, "sheet_states": {x.title: x.sheet_state for x in wb.worksheets}, "max_column": ws.max_column})
