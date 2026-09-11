from __future__ import annotations

from collections import defaultdict
from typing import Any

from models.schemas import Issue, ParseResult, PersonnelRecord
from .common import find_header_row, load_xlsx, normalize_role, normalize_text, resolve_header, source_name


def parse_personnel(source: Any, config: dict) -> ParseResult:
    filename = source_name(source, "人员信息表.xlsx")
    issues: list[Issue] = []
    try:
        wb = load_xlsx(source, data_only=True)
    except Exception as exc:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason=f"人员信息表无法打开: {exc}", suggestion="上传有效的xlsx文件")])

    aliases = config["field_aliases"]["personnel"]
    candidates = []
    department_or_role_aliases = [
        *aliases.get("department_name", []),
        *aliases.get("accounting_role", []),
    ]
    for ws in wb.worksheets:
        row = find_header_row(ws, [aliases["employee_name"], department_or_role_aliases])
        if row:
            candidates.append((ws, row))
    if len(candidates) != 1:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason="无法唯一识别人员信息工作表", suggestion="确保只有一个工作表包含姓名，以及部门名称或岗位表头")], metadata={"sheet_names": wb.sheetnames})

    ws, header_row = candidates[0]
    headers = [ws.cell(header_row, col).value for col in range(1, ws.max_column + 1)]
    try:
        name_idx, _ = resolve_header(headers, aliases["employee_name"], "员工姓名")
        dep_name_idx, _ = resolve_header(headers, aliases.get("department_name", []), "部门名称", False)
        role_idx, _ = resolve_header(headers, aliases.get("accounting_role", []), "岗位/会计归属", False)
        id_idx, _ = resolve_header(headers, aliases.get("employee_id", []), "员工编号", False)
        dep_code_idx, _ = resolve_header(headers, aliases.get("department_code", []), "部门编码", False)
        if dep_name_idx is None and role_idx is None:
            raise KeyError("部门名称或岗位")
    except (KeyError, ValueError) as exc:
        return ParseResult(issues=[Issue("E002", "错误", True, filename, ws.title, header_row, reason=f"人员信息表必要字段识别失败: {exc}", suggestion="补充唯一的姓名，以及部门名称或岗位字段")], metadata={"sheet": ws.title, "header_row": header_row})

    role_aliases = config.get("role_aliases", {})
    records: list[PersonnelRecord] = []
    for row in range(header_row + 1, ws.max_row + 1):
        raw_name = ws.cell(row, name_idx + 1).value
        if normalize_text(raw_name) == "":
            continue
        name = normalize_text(raw_name)
        raw_dep_name = ws.cell(row, dep_name_idx + 1).value if dep_name_idx is not None else None
        raw_role = ws.cell(row, role_idx + 1).value if role_idx is not None else None
        dep_name = normalize_text(raw_dep_name)
        role_text = normalize_text(raw_role)
        accounting_role = normalize_role(role_text or dep_name, role_aliases)
        if not dep_name:
            dep_name = role_text
        employee_id = normalize_text(ws.cell(row, id_idx + 1).value) if id_idx is not None else ""
        dep_code = normalize_text(ws.cell(row, dep_code_idx + 1).value) if dep_code_idx is not None else ""
        records.append(PersonnelRecord(name, accounting_role, employee_id, dep_code, dep_name, ws.title, row))
        if str(raw_name) != name:
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, employee_name=name, position=accounting_role, original_value=raw_name, reason="人员姓名空格已标准化", suggestion="无需处理；系统使用标准化姓名精确匹配"))
        if raw_dep_name is not None and str(raw_dep_name).strip() != normalize_text(raw_dep_name):
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, employee_name=name, position=accounting_role, original_value=raw_dep_name, reason="人员部门名称空格已标准化", suggestion="无需处理；系统使用标准化部门归属"))
        if raw_role is not None and str(raw_role).strip() != normalize_text(raw_role):
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, employee_name=name, position=accounting_role, original_value=raw_role, reason="人员岗位空格已标准化", suggestion="无需处理；系统使用标准化会计归属"))
        if dep_code == "006" or dep_name == "运营" or accounting_role == "运营":
            issues.append(Issue("E010", "错误", True, filename, ws.title, row, employee_name=name, position=accounting_role, original_value=f"{dep_code} {dep_name}", reason="出现已停用部门运营006", suggestion="由业务人员映射为有效部门后重新上传"))

    by_name: dict[str, list[PersonnelRecord]] = defaultdict(list)
    for rec in records:
        by_name[rec.name].append(rec)
    deduped: list[PersonnelRecord] = []
    for name, group in by_name.items():
        accounting_roles = {x.position for x in group}
        if len(accounting_roles) > 1:
            source_values = {(x.employee_id, x.department_code, x.department_name, x.position) for x in group}
            issues.append(Issue("E004", "错误", True, filename, ws.title, group[0].source_row, employee_name=name, original_value=str(sorted(source_values)), reason="人员信息表同名人员对应多个会计归属，无法唯一确定", suggestion="补充唯一人员标识或修正重复人员记录"))
        else:
            deduped.append(group[0])
            if len(group) > 1:
                issues.append(Issue("W002", "提示", False, filename, ws.title, group[0].source_row, employee_name=name, position=group[0].position, original_value=len(group), reason="人员信息表存在同名同会计归属记录，已按会计归属去重", suggestion="建议清理重复人员记录；不同会计归属仍会阻断"))
    return ParseResult(deduped, issues, {"sheet": ws.title, "header_row": header_row, "sheet_names": wb.sheetnames, "record_count": len(records)})
