from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from models.schemas import Issue, ProcessingResult
from parsers.allocation_parser import parse_allocation
from parsers.didi_parser import parse_didi
from parsers.personnel_parser import parse_personnel
from parsers.voucher_template_parser import parse_voucher_template
from .accounting import build_voucher_lines
from .allocation import allocate_orders, round_money
from .exporters import file_digest, rows_to_xlsx
from .matcher import match_orders
from .order_summary import build_order_detail_rows, summarize_order_rows, validate_order_summary
from .validation import validate_voucher
from .voucher import generate_voucher


def _serialize_detail(detail) -> dict:
    return {
        "企业订单号": detail.order_id,
        "员工姓名": detail.employee_name,
        "人员归属类别": detail.position,
        "有效价税合计": detail.gross_amount,
        "单笔不含税金额": detail.order_net_amount,
        "分摊类型": detail.allocation_type,
        "项目": detail.project_name,
        "项目比例": detail.project_ratio,
        "分摊金额（高精度）": detail.allocation_amount,
        "科目编码": detail.account_code,
        "科目名称": detail.account_name,
        "部门编码": detail.department_code,
        "部门名称": detail.department_name,
        "项目编码": detail.project_code,
        "币种": detail.currency_code,
        "来源工作表": detail.source_sheet,
        "来源行号": detail.source_row,
    }


def _serialize_voucher(line) -> dict:
    return {
        "分录类型": line.line_type,
        "科目编码": line.account_code,
        "科目名称": line.account_name,
        "借方": line.debit,
        "贷方": line.credit,
        "部门编码": line.department_code,
        "部门名称": line.department_name,
        "项目编码": line.project_code,
        "项目名称": line.project_name,
        "供应商编码": line.dimension_code,
        "供应商名称": line.dimension_name,
        "尾差调整": line.tail_adjustment,
    }


def _cleaning_log(issue: Issue) -> dict:
    return {
        "日志编码": issue.code,
        "文件名": issue.file_name,
        "工作表": issue.sheet,
        "行号": issue.row,
        "订单号": issue.order_id,
        "员工姓名": issue.employee_name,
        "岗位": issue.position,
        "原始值": issue.original_value,
        "处理结果": issue.reason,
        "建议": issue.suggestion,
    }


def _calculation_balance_rows(result: ProcessingResult) -> list[dict]:
    metrics = result.metrics
    return [
        {
            "检查项": "订单分摊明细合计=有效订单不含税金额",
            "结果": metrics["order_detail_allocation_total"] == metrics["monthly_net_amount"],
            "左值": metrics["order_detail_allocation_total"],
            "右值": metrics["monthly_net_amount"],
        },
        {
            "检查项": "岗位项目汇总合计=有效订单不含税金额",
            "结果": metrics["allocation_summary_total"] == metrics["monthly_net_amount"],
            "左值": metrics["allocation_summary_total"],
            "右值": metrics["monthly_net_amount"],
        },
    ]


def _refresh_outputs(result: ProcessingResult, *, include_voucher: bool = False) -> None:
    accounting_trace_rows = [_serialize_detail(detail) for detail in result.details]
    issue_rows = [issue.as_dict() for issue in result.issues]
    result.cleaning_logs = [_cleaning_log(issue) for issue in result.issues if not issue.blocking]
    result.output_files["处理明细.xlsx"] = rows_to_xlsx(
        {
            "订单分摊明细": result.order_detail_rows,
            "会计重分类追溯": accounting_trace_rows,
            "人员和工时匹配": result.match_rows,
            "清洗日志": result.cleaning_logs,
            "运行日志": result.audit_log,
        }
    )

    balance_rows = _calculation_balance_rows(result)
    summary_sheets = {"岗位项目分摊汇总": result.summary_rows}
    if include_voucher:
        metrics = result.metrics
        balance_rows.extend(
            [
                {"检查项": "费用借方合计=月度不含税金额", "结果": metrics["expense_after_tail"] == metrics["monthly_net_amount"], "左值": metrics["expense_after_tail"], "右值": metrics["monthly_net_amount"]},
                {"检查项": "费用借方+税额=价税合计", "结果": metrics["expense_after_tail"] + metrics["tax_amount"] == metrics["gross_amount"], "左值": metrics["expense_after_tail"] + metrics["tax_amount"], "右值": metrics["gross_amount"]},
                {"检查项": "借方合计=贷方合计", "结果": metrics["debit_total"] == metrics["credit_total"], "左值": metrics["debit_total"], "右值": metrics["credit_total"]},
            ]
        )
        summary_sheets["凭证分录预览"] = [_serialize_voucher(line) for line in result.voucher_lines]
    summary_sheets["平衡检查"] = balance_rows
    result.output_files["分摊汇总.xlsx"] = rows_to_xlsx(summary_sheets)
    result.output_files["异常清单.xlsx"] = rows_to_xlsx({"异常清单": issue_rows})


def process_calculation(personnel_source: Any, allocation_source: Any, didi_source: Any, target_month: str, config: dict) -> ProcessingResult:
    """Stage 1: calculate raw-order allocation details and summary, without creating a voucher."""
    p = parse_personnel(personnel_source, config)
    a = parse_allocation(allocation_source, target_month, config)
    d = parse_didi(didi_source, target_month, config)
    issues = [*p.issues, *a.issues, *d.issues]

    non_rnd_project = config["constants"]["non_rnd_project"]
    project_names = [name for name in config["projects"] if name != non_rnd_project]
    tolerance = Decimal(str(config["constants"]["ratio_tolerance"]))
    tax_rate = Decimal(str(config["constants"]["tax_rate"]))

    matched, match_rows, match_issues = match_orders(
        d.records,
        p.records,
        a.records,
        project_names=project_names,
        employee_role_overrides=config.get("employee_role_overrides", {}),
    )
    issues.extend(match_issues)
    details, allocation_issues = allocate_orders(matched, config)
    issues.extend(allocation_issues)
    valid_orders = [order for order in d.records if order.valid]

    order_detail_rows, detail_issues = build_order_detail_rows(matched, details, project_names, issues, tolerance, tax_rate)
    issues.extend(detail_issues)
    summary_rows = summarize_order_rows(order_detail_rows, project_names)
    summary_issues, summary_metrics = validate_order_summary(order_detail_rows, summary_rows, valid_orders, project_names, tolerance, tax_rate)
    issues.extend(summary_issues)

    gross = round_money(sum((order.enterprise_paid for order in valid_orders), Decimal("0")))
    monthly_net = round_money(sum((round_money(order.enterprise_paid / (Decimal("1") + tax_rate)) for order in valid_orders), Decimal("0")))
    unique_matched = len({order.employee_name for order, _, _ in matched})
    unmatched = sorted({row["员工姓名"] for row in match_rows if row["匹配状态"] == "失败"})
    metrics = {
        "gross_amount": gross,
        "effective_gross_amount": gross,
        "same_month_refund_amount": sum((order.refund_amount for order in d.records if order.refund_handling in {"同月部分退款按净额计算", "同月全额退款排除"}), Decimal("0")),
        "partial_refund_order_count": sum(order.refund_handling == "同月部分退款按净额计算" for order in d.records),
        "full_refund_order_count": sum(order.refund_handling == "同月全额退款排除" for order in d.records),
        "monthly_net_amount": monthly_net,
        **summary_metrics,
        "raw_order_count": len(d.records),
        "valid_order_count": len(valid_orders),
        "excluded_order_count": len(d.records) - len(valid_orders),
        "matched_people_count": unique_matched,
        "unmatched_people": unmatched,
        "blocking_issue_count": sum(issue.blocking for issue in issues),
        "calculation_ready_for_confirmation": not any(issue.blocking for issue in issues),
    }
    recognized = {
        "人员信息表": p.metadata.get("sheet", ""),
        "工时分摊表": a.metadata.get("sheet", ""),
        "滴滴账单": d.metadata.get("sheet", ""),
        "账单结算周期工作表": d.metadata.get("settlement_sheet", ""),
        "账单结算周期": d.metadata.get("settlement_value", ""),
        "账单解析月份": d.metadata.get("settlement_month", ""),
    }
    now = datetime.now().isoformat(timespec="seconds")
    audit_log = [
        {"事件": "原始账单计算开始", "时间": now, "目标月份": target_month, "规则版本": config.get("version", ""), "配置生效日期": config.get("effective_date", "")},
        {"事件": "输入指纹", "人员信息表SHA256": file_digest(personnel_source), "工时表SHA256": file_digest(allocation_source), "滴滴账单SHA256": file_digest(didi_source)},
        {"事件": "账单期间识别", "结算周期工作表": d.metadata.get("settlement_sheet", ""), "原始结算周期": d.metadata.get("settlement_value", ""), "解析月份": d.metadata.get("settlement_month", "")},
    ]
    result = ProcessingResult(
        target_month=target_month,
        issues=issues,
        details=details,
        match_rows=match_rows,
        summary_rows=summary_rows,
        voucher_lines=[],
        metrics=metrics,
        recognized_sheets=recognized,
        audit_log=audit_log,
        order_detail_rows=order_detail_rows,
        valid_orders=valid_orders,
    )
    _refresh_outputs(result)
    return result


def finalize_voucher(calculation: ProcessingResult, template_source: Any, config: dict, *, confirmed: bool) -> ProcessingResult:
    """Stage 2: after explicit manual confirmation, execute existing voucher logic."""
    calculation.output_files.pop("月度凭证.xlsx", None)
    if not confirmed:
        calculation.issues.append(Issue("E020", "错误", True, "人工确认", reason="分摊明细和汇总尚未人工确认", suggestion="核对第一阶段输出后点击确认"))
        calculation.metrics["blocking_issue_count"] = sum(issue.blocking for issue in calculation.issues)
        _refresh_outputs(calculation)
        return calculation

    calculation.audit_log.append({"事件": "人工确认分摊结果", "时间": datetime.now().isoformat(timespec="seconds"), "确认状态": "已确认"})
    parsed_template = parse_voucher_template(template_source)
    calculation.issues.extend(parsed_template.issues)
    calculation.recognized_sheets["凭证模板"] = parsed_template.metadata.get("sheet", "")
    calculation.audit_log.append({"事件": "凭证模板指纹", "凭证模板SHA256": file_digest(template_source)})

    lines, accounting_issues, voucher_metrics, audit = build_voucher_lines(calculation.details, calculation.valid_orders, config)
    calculation.voucher_lines = lines
    calculation.issues.extend(accounting_issues)
    calculation.metrics.update(voucher_metrics)
    calculation.audit_log.extend(audit)
    calculation.issues.extend(validate_voucher(lines, calculation.metrics, config))
    calculation.metrics["blocking_issue_count"] = sum(issue.blocking for issue in calculation.issues)
    calculation.metrics["can_generate_voucher"] = not any(issue.blocking for issue in calculation.issues)
    _refresh_outputs(calculation, include_voucher=True)
    if calculation.can_generate_voucher:
        calculation.output_files["月度凭证.xlsx"] = generate_voucher(template_source, parsed_template.metadata, calculation.target_month, lines, config)
    return calculation


def process_files(personnel_source: Any, allocation_source: Any, didi_source: Any, template_source: Any, target_month: str, config: dict) -> ProcessingResult:
    """Compatibility API used by tests and scripts: calculate, confirm, then finalize."""
    calculation = process_calculation(personnel_source, allocation_source, didi_source, target_month, config)
    return finalize_voucher(calculation, template_source, config, confirmed=True)


def save_outputs(result: ProcessingResult, output_dir: str | Path) -> list[Path]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for name, data in result.output_files.items():
        path = output_dir / name
        path.write_bytes(data)
        paths.append(path)
    return paths
