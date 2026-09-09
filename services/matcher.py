from __future__ import annotations

from collections import defaultdict
from decimal import Decimal
from models.schemas import AllocationRecord, DidiOrder, Issue, PersonnelRecord


def match_orders(
    orders: list[DidiOrder],
    personnel: list[PersonnelRecord],
    allocations: list[AllocationRecord],
    personnel_file: str = "人员信息表",
    allocation_file: str = "工时分摊表",
    project_names: list[str] | None = None,
    employee_role_overrides: dict[str, str] | None = None,
) -> tuple[list[tuple[DidiOrder, PersonnelRecord, AllocationRecord]], list[dict], list[Issue]]:
    issues: list[Issue] = []
    matched = []
    rows = []
    overrides = employee_role_overrides or {}
    people_by_name: dict[str, list[PersonnelRecord]] = defaultdict(list)
    for person in personnel:
        people_by_name[person.name].append(person)
    alloc_by_name: dict[str, list[AllocationRecord]] = defaultdict(list)
    for allocation in allocations:
        alloc_by_name[allocation.name].append(allocation)
    no_timesheet_logged: set[tuple[str, str]] = set()

    for order in orders:
        if not order.valid:
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": "", "人员会计类别": "", "工时岗位": "", "匹配状态": "订单无效", "说明": "未进入分摊"})
            continue
        people = people_by_name.get(order.employee_name, [])
        if not people:
            issues.append(Issue("E003", "错误", True, personnel_file, order.source_sheet, order.source_row, order.order_id, order.employee_name, reason="人员信息表找不到员工", suggestion="在人员信息表补充员工姓名和部门归属"))
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": "", "人员会计类别": "", "工时岗位": "", "匹配状态": "失败", "说明": "人员找不到"})
            continue
        if len(people) != 1:
            issues.append(Issue("E004", "错误", True, personnel_file, order.source_sheet, order.source_row, order.order_id, order.employee_name, original_value=len(people), reason="人员信息表同名人员无法唯一匹配", suggestion="修正人员信息表重复记录"))
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": "", "人员会计类别": "", "工时岗位": "", "匹配状态": "失败", "说明": "人员记录不唯一"})
            continue
        person = people[0]
        accounting_role = overrides.get(person.name, person.position)
        if not accounting_role:
            issues.append(Issue("E009", "错误", True, personnel_file, person.source_sheet, person.source_row, order.order_id, order.employee_name, original_value=person.department_name, reason="已匹配员工缺少部门归属或会计分类", suggestion="补充人员部门或受控人员会计类别后重新上传"))
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": person.department_name, "人员会计类别": "", "工时岗位": "", "匹配状态": "失败", "说明": "会计归属为空"})
            continue

        accounting_person = PersonnelRecord(
            person.name,
            accounting_role,
            person.employee_id,
            person.department_code,
            person.department_name,
            person.source_sheet,
            person.source_row,
        )
        same_name = alloc_by_name.get(person.name, [])
        candidates = [item for item in same_name if item.position == accounting_role]
        if len(candidates) > 1:
            issues.append(Issue("E006", "错误", True, allocation_file, candidates[0].source_sheet, candidates[0].source_row, order.order_id, order.employee_name, accounting_role, original_value=len(candidates), reason="工时表同一姓名加岗位出现多条记录，无法唯一匹配", suggestion="保留唯一一条姓名加岗位工时记录"))
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": person.department_name, "人员会计类别": accounting_role, "工时岗位": accounting_role, "匹配状态": "失败", "说明": "工时姓名加岗位重复"})
            continue
        if not candidates:
            allocation = AllocationRecord(
                person.name,
                accounting_role,
                {name: Decimal("0") for name in (project_names or [])},
                Decimal("1"),
                "无匹配工时记录",
                0,
                True,
            )
            log_key = (person.name, accounting_role)
            if log_key not in no_timesheet_logged:
                found_roles = sorted({item.position for item in same_name})
                original = f"会计类别={accounting_role}; 工时岗位={','.join(found_roles) if found_roles else '无'}"
                issues.append(Issue("W005", "提示", False, allocation_file, employee_name=person.name, position=accounting_role, original_value=original, reason="员工当月无姓名加岗位完全一致的研发工时，已按非研发100%计算", suggestion="无需处理；如实际参与研发，请统一人员会计类别与工时岗位"))
                no_timesheet_logged.add(log_key)
            matched.append((order, accounting_person, allocation))
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": person.department_name, "人员会计类别": accounting_role, "工时岗位": "、".join(sorted({item.position for item in same_name})) or "无", "会计归属类别": accounting_role, "匹配状态": "成功", "说明": "人员姓名匹配；无姓名加岗位完全一致的工时，按非研发100%"})
            continue

        allocation = candidates[0]
        if not allocation.valid:
            rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": person.department_name, "人员会计类别": accounting_role, "工时岗位": allocation.position, "匹配状态": "失败", "说明": "工时比例校验失败"})
            continue
        matched.append((order, accounting_person, allocation))
        rows.append({"订单号": order.order_id, "员工姓名": order.employee_name, "人员部门": person.department_name, "人员会计类别": accounting_role, "工时岗位": allocation.position, "会计归属类别": accounting_role, "匹配状态": "成功", "说明": "人员姓名匹配；姓名加岗位精确匹配工时；非研发占比空白按0" if allocation.non_rnd_missing else "人员姓名匹配；姓名加岗位精确匹配工时"})
    return matched, rows, issues
