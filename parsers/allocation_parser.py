from __future__ import annotations

from collections import Counter
from decimal import Decimal, InvalidOperation
from typing import Any

from models.schemas import AllocationRecord, Issue, ParseResult
from .common import as_decimal, find_header_row, load_xlsx, normalize_role, normalize_text, resolve_header, source_name


def parse_allocation(source: Any, target_month: str, config: dict) -> ParseResult:
    filename = source_name(source, "研发工时成本分摊表.xlsx")
    issues: list[Issue] = []
    try:
        wb = load_xlsx(source, data_only=True)
    except Exception as exc:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason=f"工时分摊表无法打开: {exc}", suggestion="上传有效的xlsx文件")])
    year, month = [int(x) for x in target_month.split("-")]
    expected_sheet = f"工时分摊-{year}.{month}"
    if expected_sheet not in wb.sheetnames:
        return ParseResult(issues=[Issue("E001", "错误", True, filename, reason=f"缺少目标月份工作表 {expected_sheet}", suggestion="补充目标月份工时工作表")], metadata={"sheet_names": wb.sheetnames})
    ws = wb[expected_sheet]
    aliases = config["field_aliases"]["allocation"]
    header_row = find_header_row(ws, [aliases["employee_name"], aliases["position"], aliases["non_rnd_ratio"]])
    if not header_row:
        return ParseResult(issues=[Issue("E002", "错误", True, filename, ws.title, reason="无法唯一识别工时表表头", suggestion="检查姓名、岗位、非研发占比表头")])
    headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
    try:
        name_idx, _ = resolve_header(headers, aliases["employee_name"], "姓名")
        pos_idx, _ = resolve_header(headers, aliases["position"], "岗位")
        non_idx, _ = resolve_header(headers, aliases["non_rnd_ratio"], "非研发占比")
        serial_idx, _ = resolve_header(headers, aliases.get("serial", []), "序号", False)
    except (KeyError, ValueError) as exc:
        return ParseResult(issues=[Issue("E002", "错误", True, filename, ws.title, header_row, reason=f"工时表必要字段识别失败: {exc}", suggestion="修正表头")])

    non_rnd_project = config["constants"]["non_rnd_project"]
    project_names = [name for name in config["projects"] if name != non_rnd_project]
    project_indices: dict[str, int | None] = {}
    historical_inactive_projects: list[str] = []
    for project in project_names:
        project_config = config["projects"].get(project, {})
        effective_from = normalize_text(project_config.get("effective_from"))
        try:
            idx, _ = resolve_header(headers, aliases.get(project, [project]), project, True)
            project_indices[project] = idx
        except (KeyError, ValueError) as exc:
            project_indices[project] = None
            if effective_from and target_month < effective_from:
                historical_inactive_projects.append(project)
                issues.append(Issue("W010", "提示", False, filename, ws.title, header_row, original_value=project, reason=f"研发项目 {project} 在目标月份尚未生效，缺少该列已按0处理", suggestion=f"无需处理；该项目自{effective_from}起启用"))
            else:
                issues.append(Issue("E002", "错误", True, filename, ws.title, header_row, original_value=project, reason=f"缺少或无法唯一识别研发项目列 {project}: {exc}", suggestion=f"在目标月份工时表中补充唯一的 {project} 列，即使全为0也必须保留"))

    known_headers: set[str] = set()
    for field in ("serial", "employee_name", "position", "non_rnd_ratio", "ignored_fields"):
        known_headers.update(normalize_text(x) for x in aliases.get(field, []))
    for project in project_names:
        known_headers.update(normalize_text(x) for x in aliases.get(project, [project]))
    unknown_headers = [normalize_text(value) for value in headers if normalize_text(value) and normalize_text(value) not in known_headers]
    for header in unknown_headers:
        issues.append(Issue("E009", "错误", True, filename, ws.title, header_row, original_value=header, reason="工时表出现未配置字段或未知研发项目", suggestion="确认该字段是否为新研发项目；如是，请先配置项目编码和规则"))

    role_aliases = config.get("role_aliases", {})
    tolerance = Decimal(str(config["constants"]["ratio_tolerance"]))
    records: list[AllocationRecord] = []
    for row in range(header_row + 1, ws.max_row + 1):
        raw_name = ws.cell(row, name_idx + 1).value
        raw_pos = ws.cell(row, pos_idx + 1).value
        serial = ws.cell(row, serial_idx + 1).value if serial_idx is not None else None
        if not normalize_text(raw_name) or not normalize_text(raw_pos):
            continue
        if serial_idx is not None and not (isinstance(serial, (int, float)) or normalize_text(serial).isdigit()):
            continue

        name = normalize_text(raw_name)
        position_text = normalize_text(raw_pos)
        position = normalize_role(raw_pos, role_aliases)
        if str(raw_name) != name:
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, employee_name=name, position=position, original_value=raw_name, reason="工时姓名空格已标准化", suggestion="无需处理；系统使用标准化姓名匹配"))
        if str(raw_pos) != position_text:
            issues.append(Issue("W001", "提示", False, filename, ws.title, row, employee_name=name, position=position, original_value=raw_pos, reason="工时岗位空格已标准化", suggestion="无需处理；系统使用标准化岗位匹配"))

        ratios: dict[str, Decimal] = {}
        blank_projects: list[str] = []
        parse_failed = False
        ratio_cells: dict[str, str] = {}
        for project, idx in project_indices.items():
            if idx is None:
                ratios[project] = Decimal("0")
                continue
            value = ws.cell(row, idx + 1).value
            ratio_cells[project] = ws.cell(row, idx + 1).coordinate
            if value is None or normalize_text(value) == "":
                ratios[project] = Decimal("0")
                blank_projects.append(project)
                issues.append(Issue("W003", "提示", False, filename, ws.title, row, employee_name=name, position=position, original_value=f"{project}=空白", reason=f"研发项目比例 {project} 空白，已按0处理", suggestion="无需处理；如实际参与该项目，请补充比例"))
                continue
            try:
                ratios[project] = as_decimal(value)
            except InvalidOperation:
                ratios[project] = Decimal("0")
                parse_failed = True
                issues.append(Issue("E002", "错误", True, filename, ws.title, row, employee_name=name, position=position, original_value=value, reason=f"项目比例 {project} 无法解析", suggestion="改为0到1之间的数字"))

        raw_non = ws.cell(row, non_idx + 1).value
        ratio_cells["非研发占比"] = ws.cell(row, non_idx + 1).coordinate
        if raw_non is None or normalize_text(raw_non) == "":
            non_rnd = Decimal("0")
            non_rnd_missing = True
            issues.append(Issue("W007", "提示", False, filename, ws.title, row, employee_name=name, position=position, original_value="空白", reason="非研发占比为空，已按0处理", suggestion="无需处理；空白表示未从事非研发项目"))
        else:
            non_rnd_missing = False
            try:
                non_rnd = as_decimal(raw_non)
            except InvalidOperation:
                non_rnd = Decimal("0")
                parse_failed = True
                issues.append(Issue("E002", "错误", True, filename, ws.title, row, employee_name=name, position=position, original_value=raw_non, reason="非研发占比无法解析", suggestion="改为0到1之间的数字"))

        valid = not parse_failed
        all_ratios = {**ratios, "非研发占比": non_rnd}
        out_of_range = [(key, value) for key, value in all_ratios.items() if value < 0 or value > 1]
        if out_of_range:
            valid = False
            issues.append(Issue("E007", "错误", True, filename, ws.title, row, employee_name=name, position=position, original_value=str(out_of_range), reason="项目比例或非研发占比超出0到1", suggestion="修正比例"))

        rnd_sum = sum(ratios.values(), Decimal("0"))
        expected_non = Decimal("1") - rnd_sum
        total = rnd_sum + non_rnd
        if abs(non_rnd - expected_non) > tolerance or abs(total - Decimal("1")) > tolerance:
            valid = False
            detail = {
                "各项目比例": ratios,
                "非研发占比": non_rnd,
                "研发比例合计": rnd_sum,
                "理论非研发占比": expected_non,
                "全部比例合计": total,
                "来源单元格": ratio_cells,
            }
            issues.append(Issue("E008", "错误", True, filename, ws.title, row, employee_name=name, position=position, original_value=str(detail), reason="研发项目比例与非研发占比合计不为1或交叉校验不一致", suggestion="按实际工时修正比例，合计必须为1"))
        records.append(AllocationRecord(name, position, ratios, non_rnd, ws.title, row, valid, non_rnd_missing, tuple(blank_projects)))

    counts = Counter((x.name, x.position) for x in records)
    for (name, position), count in counts.items():
        if count > 1:
            issues.append(Issue("E006", "错误", True, filename, ws.title, employee_name=name, position=position, original_value=count, reason="工时表同一姓名加岗位出现多条记录", suggestion="保留唯一一条姓名加岗位记录，不得求和或取第一条"))
            for rec in records:
                if rec.name == name and rec.position == position:
                    rec.valid = False

    return ParseResult(
        records,
        issues,
        {
            "sheet": ws.title,
            "header_row": header_row,
            "sheet_names": wb.sheetnames,
            "record_count": len(records),
            "missing_projects": [key for key, value in project_indices.items() if value is None and key not in historical_inactive_projects],
            "historical_inactive_projects": historical_inactive_projects,
            "unknown_headers": unknown_headers,
        },
    )
