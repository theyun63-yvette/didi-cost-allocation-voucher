from __future__ import annotations

from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
from typing import Any

from models.schemas import DidiOrder, Issue, ParseResult
from .common import as_decimal, find_header_row, load_xlsx, normalize_text, resolve_header, source_name

CANCEL_STATUS_WORDS = ("取消", "关闭")
RED_STATUS_WORDS = ("红字",)
REFUND_STATUS_WORDS = ("退款",)


def _parse_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = normalize_text(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def _month_from_value(value: Any) -> str | None:
    if isinstance(value, (datetime, date)):
        return f"{value.year:04d}-{value.month:02d}"
    text = normalize_text(value)
    if not text:
        return None
    patterns = (
        r"^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?$",
        r"^(\d{4})年(\d{1,2})月(?:\d{1,2}日)?$",
        r"^(\d{4})(\d{2})$",
    )
    for pattern in patterns:
        match = re.match(pattern, text)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            if 1 <= month <= 12:
                return f"{year:04d}-{month:02d}"
    return None


def _parse_settlement_range(value: Any) -> tuple[str | None, str | None]:
    """Return (month, error) for a monthly settlement period."""
    if isinstance(value, (datetime, date)):
        month = f"{value.year:04d}-{value.month:02d}"
        return month, None
    text = normalize_text(value)
    if not text:
        return None, "结算周期为空"
    dates = re.findall(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?", text)
    if dates:
        months = {f"{int(y):04d}-{int(m):02d}" for y, m, _ in dates}
        if len(months) == 1:
            return months.pop(), None
        return None, f"结算周期跨月或包含多个期间: {text}"
    month = _month_from_value(text)
    if month:
        return month, None
    return None, f"结算周期无法解析: {text}"


def _find_settlement_period(wb, aliases: list[str], summary_sheet_aliases: list[str] | None = None) -> tuple[str, str, int, Any] | None:
    alias_set = {normalize_text(x) for x in aliases}
    preferred_sheets = {normalize_text(x) for x in (summary_sheet_aliases or [])}
    matches: list[tuple[str, str, int, Any]] = []
    worksheets = [ws for ws in wb.worksheets if normalize_text(ws.title) in preferred_sheets]
    if not worksheets:
        worksheets = list(wb.worksheets)
    for ws in worksheets:
        for row in range(1, min(ws.max_row, 50) + 1):
            for col in range(1, min(ws.max_column, 30) + 1):
                if normalize_text(ws.cell(row, col).value) not in alias_set:
                    continue
                value = None
                for offset in range(1, 5):
                    candidate = ws.cell(row, col + offset).value
                    if normalize_text(candidate):
                        value = candidate
                        break
                matches.append((ws.title, ws.cell(row, col).coordinate, row, value))
    return matches[0] if len(matches) == 1 else None


def _resolve_preferred_header(headers: list[Any], aliases: list[str]) -> tuple[int | None, str | None]:
    normalized_headers = [normalize_text(value) for value in headers]
    for alias in aliases:
        matches = [index for index, header in enumerate(normalized_headers) if header == normalize_text(alias)]
        if len(matches) > 1:
            raise ValueError(f"订单日期候选字段“{alias}”重复")
        if matches:
            return matches[0], alias
    return None, None


def parse_didi(source: Any, target_month: str, config: dict) -> ParseResult:
    filename = source_name(source, "滴滴网约车账单.xlsx")
    issues: list[Issue] = []
    try:
        wb = load_xlsx(source, data_only=True)
    except Exception as exc:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason=f"滴滴账单无法打开: {exc}", suggestion="上传有效的xlsx文件")])

    aliases = config["field_aliases"]["didi"]
    settlement = _find_settlement_period(
        wb,
        aliases.get("settlement_period", ["结算周期"]),
        aliases.get("summary_sheet", ["汇总数据"]),
    )
    settlement_sheet = ""
    settlement_value: Any = ""
    settlement_month: str | None = None
    if settlement is None:
        issues.append(Issue("E012", "错误", True, filename, reason="无法唯一识别账单结算周期", suggestion="确保汇总页只有一个结算周期字段并填写明确的起止日期"))
    else:
        settlement_sheet, settlement_cell, settlement_row, settlement_value = settlement
        settlement_month, error = _parse_settlement_range(settlement_value)
        if error:
            issues.append(Issue("E012", "错误", True, filename, settlement_sheet, settlement_row, original_value=settlement_value, reason=error, suggestion="修正账单结算周期为同一自然月"))
        elif settlement_month != target_month:
            issues.append(Issue("E012", "错误", True, filename, settlement_sheet, settlement_row, original_value=settlement_value, reason=f"账单结算周期{settlement_month}与目标月份{target_month}不一致", suggestion="选择正确月份或上传对应月份账单"))

    candidates = []
    for ws in wb.worksheets:
        row = find_header_row(
            ws,
            [
                aliases["order_id"],
                aliases["billing_period"],
                aliases["employee_name"],
                aliases["enterprise_paid"],
                aliases["refund_amount"],
            ],
            max_rows=20,
        )
        if row:
            candidates.append((ws, row))
    if len(candidates) != 1:
        return ParseResult(
            issues=[
                *issues,
                Issue("E001", "错误", True, filename, reason="无法唯一识别滴滴用车订单工作表", suggestion="确保只有一个工作表包含所属账期、订单号、下单人、企业实付和企业实退表头"),
            ],
            metadata={"sheet_names": wb.sheetnames, "settlement_month": settlement_month or "", "settlement_value": settlement_value},
        )

    ws, header_row = candidates[0]
    headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
    try:
        order_idx, _ = resolve_header(headers, aliases["order_id"], "企业订单号")
        period_idx, _ = resolve_header(headers, aliases["billing_period"], "所属账期")
        name_idx, _ = resolve_header(headers, aliases["employee_name"], "下单人姓名")
        paid_idx, _ = resolve_header(headers, aliases["enterprise_paid"], "企业实付金额")
        refund_idx, _ = resolve_header(headers, aliases["refund_amount"], "企业实退金额")
        payable_idx, _ = resolve_header(headers, aliases.get("enterprise_payable", []), "企业应付金额", False)
        date_idx, date_header = _resolve_preferred_header(headers, aliases.get("order_date", []))
        status_idx, _ = resolve_header(headers, aliases.get("order_status", []), "订单状态", False)
        cancel_idx, _ = resolve_header(headers, aliases.get("cancel_time", []), "取消时间", False)
        refund_time_idx, _ = resolve_header(headers, aliases.get("refund_time", []), "退款时间", False)
    except (KeyError, ValueError) as exc:
        return ParseResult(issues=[*issues, Issue("E002", "错误", True, filename, ws.title, header_row, reason=f"滴滴账单必要字段识别失败: {exc}", suggestion="修正表头或使用标准账单")])

    current_period_values = {normalize_text(x) for x in config["constants"].get("current_period_values", ["本期"])}
    orders: list[DidiOrder] = []
    for row in range(header_row + 1, ws.max_row + 1):
        oid = normalize_text(ws.cell(row, order_idx + 1).value)
        if not oid:
            continue

        raw_name_value = ws.cell(row, name_idx + 1).value
        raw_name = "" if raw_name_value is None else str(raw_name_value)
        name = normalize_text(raw_name_value)
        if raw_name != name:
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, oid, name, original_value=raw_name, reason="下单人姓名空格已标准化", suggestion="无需处理；系统使用标准化姓名精确匹配"))
        if not name:
            issues.append(Issue("E003", "错误", True, filename, ws.title, row, oid, original_value=raw_name, reason="下单人姓名为空", suggestion="补充下单人姓名后重新上传"))

        raw_period = ws.cell(row, period_idx + 1).value
        period_text = normalize_text(raw_period)
        if period_text in current_period_values:
            order_month = settlement_month
        else:
            order_month = _month_from_value(raw_period)
        valid = True
        if not period_text:
            issues.append(Issue("E012", "错误", True, filename, ws.title, row, oid, name, original_value=raw_period, reason="所属账期为空", suggestion="补充可识别的所属账期"))
            valid = False
        elif not order_month:
            issues.append(Issue("E012", "错误", True, filename, ws.title, row, oid, name, original_value=raw_period, reason="所属账期无法解析", suggestion="填写明确年月，或使用可由汇总页结算周期解释的本期标识"))
            valid = False
        elif order_month != target_month:
            issues.append(Issue("W006", "提示", False, filename, ws.title, row, oid, name, original_value=raw_period, reason=f"订单所属账期{order_month}不属于目标月份{target_month}，已排除", suggestion="无需处理；如应计入本月，请修正所属账期"))
            valid = False

        raw_paid = ws.cell(row, paid_idx + 1).value
        raw_refund = ws.cell(row, refund_idx + 1).value
        raw_payable = ws.cell(row, payable_idx + 1).value if payable_idx is not None else None
        status = normalize_text(ws.cell(row, status_idx + 1).value) if status_idx is not None else ""
        payable: Decimal | None = None
        try:
            original_paid = as_decimal(raw_paid)
            refund = as_decimal(raw_refund, blank_zero=True)
            if normalize_text(raw_payable):
                payable = as_decimal(raw_payable)
        except InvalidOperation:
            issues.append(Issue("E002", "错误", True, filename, ws.title, row, oid, name, original_value=f"实付={raw_paid}, 实退={raw_refund}, 应付={raw_payable}", reason="订单金额无法解析", suggestion="修正金额格式"))
            original_paid, refund, valid = Decimal("0"), Decimal("0"), False

        effective_paid = original_paid
        refund_handling = "无退款"
        raw_date = ws.cell(row, date_idx + 1).value if date_idx is not None else None
        order_dt = _parse_datetime(raw_date) if raw_date not in (None, "") else None
        cancel_time = normalize_text(ws.cell(row, cancel_idx + 1).value) if cancel_idx is not None else ""
        raw_refund_time = ws.cell(row, refund_time_idx + 1).value if refund_time_idx is not None else None
        refund_time_text = normalize_text(raw_refund_time)
        refund_dt = _parse_datetime(raw_refund_time) if refund_time_text else None
        reconciliation_tolerance = Decimal(str(config["constants"].get("refund_reconciliation_tolerance", "0.01")))

        if original_paid < 0 or refund < 0:
            issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}", reason="企业实付金额或企业实退金额为负数", suggestion="负数及红字订单需单独处理"))
            valid = False
        elif refund > original_paid:
            issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}", reason="企业实退金额大于企业实付金额", suggestion="核对原订单及退款数据后重新上传"))
            valid = False
        elif refund > 0:
            effective_paid = original_paid - refund
            refund_handling = "待校验退款"
            if not refund_time_text:
                issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}, 退款时间为空", reason="退款月份无法确认", suggestion="补充退款时间；跨月退款需单独处理"))
                valid = False
            elif refund_dt is None:
                issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=refund_time_text, reason="退款时间无法解析", suggestion="使用明确的年月日时间；跨月退款需单独处理"))
                valid = False
            elif f"{refund_dt.year:04d}-{refund_dt.month:02d}" != target_month:
                issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"订单月份={order_month}, 退款时间={refund_time_text}", reason="检测到跨月退款", suggestion="跨月退款及红字处理需单独确认，不得直接冲减本月普通订单"))
                valid = False
            elif payable is not None and abs(payable - effective_paid) > reconciliation_tolerance:
                issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}, 实付减实退={effective_paid}, 企业应付={payable}", reason="退款净额与企业应付金额不一致", suggestion="核对企业实付、实退和应付金额后重新上传"))
                valid = False
            elif effective_paid == 0:
                refund_handling = "同月全额退款排除"
                issues.append(Issue("W009", "提示", False, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}, 退款时间={refund_time_text}", reason="同月全额退款，净额为0，订单已排除", suggestion="无需处理；该订单仅保留审计记录"))
                valid = False
            else:
                refund_handling = "同月部分退款按净额计算"
                issues.append(Issue("W008", "提示", False, filename, ws.title, row, oid, name, original_value=f"实付={original_paid}, 实退={refund}, 有效价税合计={effective_paid}, 退款时间={refund_time_text}", reason="同月部分退款，已按企业实付减企业实退的净额计算", suggestion="请在第一阶段明细中复核退款净额"))
        elif refund_time_text or any(word in status for word in REFUND_STATUS_WORDS):
            issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"状态={status}, 实退={refund}, 退款时间={refund_time_text}", reason="订单显示退款状态或退款时间，但企业实退金额为0", suggestion="核对退款金额后重新上传"))
            valid = False
        elif original_paid == 0:
            issues.append(Issue("W004", "提示", False, filename, ws.title, row, oid, name, original_value=original_paid, reason="企业实付金额为0，订单已排除", suggestion="核实后可保留为审计记录"))
            valid = False

        if cancel_time or any(word in status for word in CANCEL_STATUS_WORDS):
            issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"状态={status}, 取消时间={cancel_time}", reason="取消或关闭订单", suggestion="当前版本不自动处理取消或关闭订单，请核对后重新上传"))
            valid = False
        if any(word in status for word in RED_STATUS_WORDS):
            issues.append(Issue("E013", "错误", True, filename, ws.title, row, oid, name, original_value=f"状态={status}", reason="检测到红字订单", suggestion="红字订单需单独处理"))
            valid = False

        raw = {normalize_text(h): ws.cell(row, i + 1).value for i, h in enumerate(headers) if normalize_text(h)}
        orders.append(
            DidiOrder(
                oid,
                name,
                effective_paid,
                refund,
                order_dt,
                ws.title,
                row,
                status,
                raw,
                valid and bool(name),
                raw_employee_name=raw_name,
                billing_period=period_text,
                billing_month=order_month or "",
                original_enterprise_paid=original_paid,
                enterprise_payable=payable,
                refund_datetime=refund_dt,
                refund_handling=refund_handling,
            )
        )

    counts = Counter(x.order_id for x in orders)
    for oid, count in counts.items():
        if count > 1:
            issues.append(Issue("E011", "错误", True, filename, ws.title, order_id=oid, original_value=count, reason="企业订单号重复", suggestion="去重并确认唯一订单后重新上传"))
            for order in orders:
                if order.order_id == oid:
                    order.valid = False

    return ParseResult(
        orders,
        issues,
        {
            "sheet": ws.title,
            "header_row": header_row,
            "sheet_names": wb.sheetnames,
            "record_count": len(orders),
            "date_field": date_header or "",
            "billing_period_field": normalize_text(headers[period_idx]),
            "settlement_sheet": settlement_sheet,
            "settlement_value": settlement_value,
            "settlement_month": settlement_month or "",
        },
    )
