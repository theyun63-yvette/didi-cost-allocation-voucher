from __future__ import annotations

from decimal import Decimal
from models.schemas import Issue, VoucherLine
from .allocation import round_money


def validate_voucher(lines: list[VoucherLine], metrics: dict, config: dict) -> list[Issue]:
    issues: list[Issue] = []
    expense = round_money(sum((x.debit for x in lines if x.line_type == "费用"), Decimal("0")))
    debit = round_money(sum((x.debit for x in lines), Decimal("0")))
    credit = round_money(sum((x.credit for x in lines), Decimal("0")))
    ap_credit = round_money(sum((x.credit for x in lines if x.account_code == config["accounts"]["accounts_payable"]["code"]), Decimal("0")))
    gross = round_money(metrics["gross_amount"])
    monthly_net = round_money(metrics["monthly_net_amount"])
    tax = round_money(metrics["tax_amount"])
    failures = []
    if expense != monthly_net: failures.append(f"费用借方{expense} != 月度不含税{monthly_net}")
    if expense + tax != gross: failures.append(f"费用{expense}+税额{tax} != 价税合计{gross}")
    if debit != credit: failures.append(f"借方{debit} != 贷方{credit}")
    if ap_credit != gross: failures.append(f"应付账款贷方{ap_credit} != 价税合计{gross}")
    if failures:
        issues.append(Issue("E015", "错误", True, "计算结果", original_value="; ".join(failures), reason="借贷不平或费用合计不等于月度不含税金额", suggestion="检查匹配、比例、尾差和会计映射"))
    for line in lines:
        if line.department_code == "006":
            issues.append(Issue("E010", "错误", True, "凭证预览", original_value=line.as_dict(), reason="凭证分录出现停用部门006", suggestion="修正源数据或配置"))
    return issues
