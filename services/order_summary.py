from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

from models.schemas import AllocationDetail, AllocationRecord, DidiOrder, Issue, PersonnelRecord
from .allocation import round_money

MatchedOrder = tuple[DidiOrder, PersonnelRecord, AllocationRecord]


def build_order_detail_rows(
    matched: list[MatchedOrder],
    details: list[AllocationDetail],
    project_names: list[str],
    existing_issues: list[Issue],
    tolerance: Decimal,
    tax_rate: Decimal,
) -> tuple[list[dict], list[Issue]]:
    """Build one deterministic calculation row per matched order."""
    details_by_order: dict[str, list[AllocationDetail]] = defaultdict(list)
    for detail in details:
        details_by_order[detail.order_id].append(detail)

    issues: list[Issue] = []
    rows: list[dict] = []
    for order, person, allocation in matched:
        order_details = details_by_order.get(order.order_id, [])
        amounts = {name: Decimal("0") for name in project_names}
        non_rnd_amount = Decimal("0")
        for detail in order_details:
            if detail.allocation_type == "研发":
                if detail.project_name in amounts:
                    amounts[detail.project_name] += detail.allocation_amount
            else:
                non_rnd_amount += detail.allocation_amount

        order_net = order_details[0].order_net_amount if order_details else round_money(
            order.enterprise_paid / (Decimal("1") + tax_rate)
        )
        allocation_total = sum(amounts.values(), Decimal("0")) + non_rnd_amount
        difference = allocation_total - order_net
        if abs(difference) > tolerance:
            issue = Issue(
                "E017",
                "错误",
                True,
                "计算结果",
                order.source_sheet,
                order.source_row,
                order.order_id,
                order.employee_name,
                person.position,
                original_value=f"单笔不含税={order_net}, 分摊合计={allocation_total}, 差额={difference}",
                reason="单笔订单分摊金额不平",
                suggestion="检查人员匹配、工时比例和项目配置",
            )
            issues.append(issue)

        related_codes = {
            issue.code
            for issue in [*existing_issues, *issues]
            if issue.order_id == order.order_id
            or (issue.employee_name == order.employee_name and (not issue.position or issue.position == person.position))
        }
        row = {
            "企业订单号": order.order_id,
            "所属账期": order.billing_period,
            "解析账期": order.billing_month,
            "下单人原始姓名": order.raw_employee_name or order.employee_name,
            "标准化姓名": order.employee_name,
            "会计归属类别/岗位": person.position,
            "企业实付金额（原始）": order.original_enterprise_paid if order.original_enterprise_paid is not None else order.enterprise_paid,
            "企业实退金额": order.refund_amount,
            "企业应付金额（账单）": order.enterprise_payable if order.enterprise_payable is not None else "",
            "退款处理方式": order.refund_handling,
            "有效价税合计": order.enterprise_paid,
            "单笔不含税金额": order_net,
        }
        for project in project_names:
            row[f"{project}比例"] = allocation.project_ratios.get(project, Decimal("0"))
        row["非研发占比"] = allocation.non_rnd_ratio
        for project in project_names:
            row[f"{project}金额"] = amounts[project]
        row["非研发金额"] = non_rnd_amount
        row["分摊金额合计"] = allocation_total
        row["单笔分摊差额"] = difference
        row.update(
            {
                "人员信息来源工作表": person.source_sheet,
                "人员信息来源行": person.source_row,
                "工时记录来源工作表": allocation.source_sheet,
                "工时记录来源行": allocation.source_row or "无",
                "账单来源工作表": order.source_sheet,
                "账单来源行": order.source_row,
                "匹配结果": "成功",
                "清洗日志或异常编号": "、".join(sorted(related_codes)),
            }
        )
        rows.append(row)
    return rows, issues


def summarize_order_rows(order_rows: list[dict], project_names: list[str]) -> list[dict]:
    """Return an Excel-pivot-equivalent wide summary by accounting role."""
    amount_fields = [f"{name}金额" for name in project_names] + ["非研发金额"]
    grouped: dict[str, dict[str, Decimal]] = defaultdict(lambda: {field: Decimal("0") for field in amount_fields})
    for row in order_rows:
        role = str(row["会计归属类别/岗位"])
        for field in amount_fields:
            grouped[role][field] += row[field]

    rows: list[dict] = []
    for role in sorted(grouped):
        values = grouped[role]
        rows.append(
            {
                "会计归属类别/岗位": role,
                **values,
                "合计金额": sum(values.values(), Decimal("0")),
            }
        )
    if rows:
        total_values = {
            field: sum((row[field] for row in rows), Decimal("0"))
            for field in amount_fields
        }
        rows.append(
            {
                "会计归属类别/岗位": "合计",
                **total_values,
                "合计金额": sum(total_values.values(), Decimal("0")),
            }
        )
    return rows


def validate_order_summary(
    order_rows: list[dict],
    summary_rows: list[dict],
    valid_orders: list[DidiOrder],
    project_names: list[str],
    tolerance: Decimal,
    tax_rate: Decimal,
) -> tuple[list[Issue], dict[str, Decimal]]:
    issues: list[Issue] = []
    expected_net = sum(
        (round_money(order.enterprise_paid / (Decimal("1") + tax_rate)) for order in valid_orders),
        Decimal("0"),
    )
    detail_net = sum((row["单笔不含税金额"] for row in order_rows), Decimal("0"))
    detail_allocated = sum((row["分摊金额合计"] for row in order_rows), Decimal("0"))
    total_row = next((row for row in summary_rows if row["会计归属类别/岗位"] == "合计"), None)
    summary_total = total_row["合计金额"] if total_row else Decimal("0")

    if abs(detail_net - expected_net) > tolerance:
        issues.append(Issue("E018", "错误", True, "计算结果", original_value=f"有效订单不含税={expected_net}, 已计算订单不含税={detail_net}", reason="订单计算明细未覆盖全部有效订单", suggestion="检查人员及工时匹配异常"))
    if abs(detail_allocated - expected_net) > tolerance:
        issues.append(Issue("E018", "错误", True, "计算结果", original_value=f"有效订单不含税={expected_net}, 分摊明细合计={detail_allocated}", reason="订单分摊明细合计与有效订单不含税金额不一致", suggestion="检查单笔分摊、比例和项目映射"))
    if abs(summary_total - expected_net) > tolerance:
        issues.append(Issue("E019", "错误", True, "计算结果", original_value=f"有效订单不含税={expected_net}, 汇总合计={summary_total}", reason="分摊汇总金额与有效订单不含税金额不一致", suggestion="检查汇总分组和金额字段"))

    project_totals = {
        name: sum((row[f"{name}金额"] for row in order_rows), Decimal("0"))
        for name in project_names
    }
    project_totals["非研发"] = sum((row["非研发金额"] for row in order_rows), Decimal("0"))
    return issues, {
        "order_detail_net_total": detail_net,
        "order_detail_allocation_total": detail_allocated,
        "allocation_summary_total": summary_total,
        **{f"project_total_{key}": value for key, value in project_totals.items()},
    }
