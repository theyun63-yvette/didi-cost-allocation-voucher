from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from models.schemas import AllocationDetail, AllocationRecord, DidiOrder, Issue, PersonnelRecord

CENT = Decimal("0.01")


def round_money(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def allocate_orders(matched: list[tuple[DidiOrder, PersonnelRecord, AllocationRecord]], config: dict) -> tuple[list[AllocationDetail], list[Issue]]:
    details: list[AllocationDetail] = []
    issues: list[Issue] = []
    tolerance = Decimal(str(config["constants"]["ratio_tolerance"]))
    tax_factor = Decimal("1") + Decimal(str(config["constants"]["tax_rate"]))
    role_rules = config["role_rules"]
    accounts = config["accounts"]
    departments = config["departments"]
    projects = config["projects"]
    currency = config["constants"]["currency_code"]

    for order, person, allocation in matched:
        rule = role_rules.get(person.position)
        if not rule:
            issues.append(Issue("E009", "错误", True, "config/rules.yaml", allocation.source_sheet, allocation.source_row, order.order_id, order.employee_name, person.position, original_value=person.position, reason="岗位会计映射缺失", suggestion="在rules.yaml补充岗位入账规则"))
            continue
        order_net = round_money(order.enterprise_paid / tax_factor)
        allocated_total = Decimal("0")
        for project_name, ratio in allocation.project_ratios.items():
            if ratio == 0:
                continue
            if project_name not in projects or projects[project_name].get("status") != "enabled":
                issues.append(Issue("E009", "错误", True, "config/rules.yaml", allocation.source_sheet, allocation.source_row, order.order_id, order.employee_name, person.position, original_value=project_name, reason="项目编码缺失、未知或已停用", suggestion="补充有效项目配置"))
                continue
            mapping = rule["rnd"]
            account = accounts.get(mapping["account"])
            department = departments.get(mapping["department"])
            project = projects[project_name]
            if not account or not department or department.get("status") != "enabled":
                code = "E010" if department and department.get("code") == "006" else "E009"
                issues.append(Issue(code, "错误", True, "config/rules.yaml", allocation.source_sheet, allocation.source_row, order.order_id, order.employee_name, person.position, original_value=str(mapping), reason="研发会计科目或部门映射无效", suggestion="修正配置，不得使用部门006"))
                continue
            amount = order_net * ratio
            allocated_total += amount
            details.append(AllocationDetail(order.order_id, order.employee_name, person.position, order.enterprise_paid, order_net, "研发", project_name, ratio, amount, account["code"], account["name"], department["code"], department["name"], project["code"], currency, order.source_sheet, order.source_row))
        if allocation.non_rnd_ratio != 0:
            mapping = rule["non_rnd"]
            account = accounts.get(mapping["account"])
            department = departments.get(mapping["department"])
            project = projects.get(mapping.get("project", config["constants"]["non_rnd_project"]))
            if not account or not department or not project or department.get("status") != "enabled":
                code = "E010" if department and department.get("code") == "006" else "E009"
                issues.append(Issue(code, "错误", True, "config/rules.yaml", allocation.source_sheet, allocation.source_row, order.order_id, order.employee_name, person.position, original_value=str(mapping), reason="非研发会计科目、部门或项目映射无效", suggestion="修正配置，不得使用部门006"))
            else:
                amount = order_net * allocation.non_rnd_ratio
                allocated_total += amount
                details.append(AllocationDetail(order.order_id, order.employee_name, person.position, order.enterprise_paid, order_net, "非研发", project["name"], allocation.non_rnd_ratio, amount, account["code"], account["name"], department["code"], department["name"], project["code"], currency, order.source_sheet, order.source_row))
        if abs(allocated_total - order_net) > tolerance:
            issues.append(Issue("E008", "错误", True, allocation.source_sheet, allocation.source_sheet, allocation.source_row, order.order_id, order.employee_name, person.position, original_value=f"不含税={order_net}, 分摊={allocated_total}", reason="单笔订单分摊不平衡", suggestion="检查工时比例和精度"))
    return details, issues
