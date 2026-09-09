from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from models.schemas import AllocationDetail, DidiOrder, Issue, VoucherLine
from .allocation import round_money


def build_voucher_lines(details: list[AllocationDetail], valid_orders: list[DidiOrder], config: dict) -> tuple[list[VoucherLine], list[Issue], dict, list[dict]]:
    issues: list[Issue] = []
    audit: list[dict] = []
    grouped: dict[tuple[str, str, str, str], dict] = {}
    for d in details:
        key = (d.account_code, d.department_code, d.project_code, d.currency_code)
        if key not in grouped:
            grouped[key] = {"amount": Decimal("0"), "sample": d}
        grouped[key]["amount"] += d.allocation_amount
    expense_lines: list[VoucherLine] = []
    for key in sorted(grouped):
        item = grouped[key]
        d = item["sample"]
        amount = round_money(item["amount"])
        if amount:
            expense_lines.append(VoucherLine("费用", d.account_code, d.account_name, debit=amount, department_code=d.department_code, department_name=d.department_name, project_code=d.project_code, project_name=d.project_name))

    gross = sum((o.enterprise_paid for o in valid_orders), Decimal("0"))
    gross = round_money(gross)
    tax_rate = Decimal(str(config["constants"]["tax_rate"]))
    tax_factor = Decimal("1") + tax_rate
    # 与订单明细保持同一口径：先逐笔计算并舍入不含税金额，再汇总月度不含税金额。
    monthly_net = sum((round_money(order.enterprise_paid / tax_factor) for order in valid_orders), Decimal("0.00"))
    monthly_net = round_money(monthly_net)
    tax = gross - monthly_net
    expense_before_tail = sum((x.debit for x in expense_lines), Decimal("0"))
    tail = monthly_net - expense_before_tail
    limit = Decimal(str(config["constants"]["tail_auto_limit_exclusive"]))
    if tail != 0:
        if abs(tail) < limit:
            tr = config["tail_rule"]
            account = config["accounts"][tr["account"]]
            department = config["departments"][tr["department"]]
            project = config["projects"][tr["project"]]
            target = next((x for x in expense_lines if (x.account_code, x.department_code, x.project_code) == (account["code"], department["code"], project["code"])), None)
            before = target.debit if target else Decimal("0")
            if target:
                target.debit += tail
                target.tail_adjustment = True
            else:
                target = VoucherLine("费用", account["code"], account["name"], debit=tail, department_code=department["code"], department_name=department["name"], project_code=project["code"], project_name=project["name"], tail_adjustment=True)
                expense_lines.append(target)
            audit.append({"事件": "尾差调整", "调整前金额": before, "尾差": tail, "调整后金额": target.debit, "科目": account["code"], "部门": department["code"], "项目": project["code"]})
        else:
            issues.append(Issue("E014", "错误", True, "计算结果", reason=f"尾差绝对值大于等于1元: {tail}", original_value=tail, suggestion="检查未匹配订单、比例及项目列后重新计算"))

    expense_lines.sort(key=lambda x: (x.account_code, x.department_code, x.project_code))
    lines = list(expense_lines)
    tax_account = config["accounts"]["input_tax_pending"]
    ap_account = config["accounts"]["accounts_payable"]
    if tax != 0:
        lines.append(VoucherLine("税额", tax_account["code"], tax_account["name"], debit=tax))
    lines.append(VoucherLine("应付账款", ap_account["code"], ap_account["name"], credit=gross, dimension_type="供应商", dimension_code=config["constants"]["supplier_code"], dimension_name=config["constants"]["supplier_name"]))
    metrics = {"gross_amount": gross, "monthly_net_amount": monthly_net, "tax_amount": tax, "expense_before_tail": expense_before_tail, "tail_difference": tail, "expense_after_tail": sum((x.debit for x in expense_lines), Decimal("0")), "debit_total": sum((x.debit for x in lines), Decimal("0")), "credit_total": sum((x.credit for x in lines), Decimal("0"))}
    return lines, issues, metrics, audit


def summarize_details(details: list[AllocationDetail]) -> list[dict]:
    grouped: dict[tuple[str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    for d in details:
        label = d.project_name if d.allocation_type == "研发" else "非研发"
        grouped[(d.position, label)] += d.allocation_amount
    return [{"岗位": k[0], "项目或类型": k[1], "分摊金额": round_money(v)} for k, v in sorted(grouped.items())]
