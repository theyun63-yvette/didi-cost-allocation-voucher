from __future__ import annotations

from copy import copy
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA_ROOT = ROOT.parent
MANUAL = DATA_ROOT / "滴滴7月网约车.xlsx"
OLD_VOUCHER = Path.home() / "Downloads" / "月度凭证-2.xlsx"
NEW_VOUCHER = ROOT / "output" / "sample-2026-07" / "月度凭证.xlsx"
OUTPUT = ROOT / "output" / "sample-2026-07" / "计算逻辑核对.xlsx"

CENT = Decimal("0.01")
BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
LIGHT_YELLOW = "FFF2CC"
LIGHT_GREEN = "E2F0D9"
LIGHT_RED = "FCE4D6"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="B7B7B7")


def d(value) -> Decimal:
    return Decimal(str(value or 0))


def q(value: Decimal) -> Decimal:
    return value.quantize(CENT, rounding=ROUND_HALF_UP)


def voucher_map(path: Path) -> dict[tuple[str, str, str], Decimal]:
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    result = {}
    for row in range(5, ws.max_row + 1):
        account = str(ws.cell(row, 8).value or "")
        department = str(ws.cell(row, 15).value or "")
        project = str(ws.cell(row, 18).value or "")
        # 旧凭证使用99999“无”，差异表中归一化到新项目00009便于金额对比。
        if project == "99999":
            project = "00009"
        debit = ws.cell(row, 10).value
        if account and debit is not None:
            result[(account, department, project)] = d(debit)
    return result


def style_table(ws, header_row: int, min_col: int, max_col: int, max_row: int) -> None:
    for cell in ws[header_row][min_col - 1:max_col]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(color=WHITE, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=header_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def add_title(ws, title: str, subtitle: str, end_col: int) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    ws["A1"] = title
    ws["A1"].font = Font(size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
    ws["A2"] = subtitle
    ws["A2"].font = Font(color="666666", italic=True)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[2].height = 32


def build() -> Path:
    manual_wb = load_workbook(MANUAL, data_only=True)
    pivot = manual_wb["Sheet1"]
    categories = []
    for row in range(4, 13):
        category = pivot.cell(row, 1).value
        if category and category != "总计":
            categories.append([category] + [d(pivot.cell(row, col).value) for col in range(2, 10)])

    old = voucher_map(OLD_VOUCHER)
    new = voucher_map(NEW_VOUCHER)
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    # 公式说明
    ws = wb.active
    ws.title = "公式说明"
    add_title(ws, "2026年7月滴滴车费计算公式", "金额采用 Decimal 和 ROUND_HALF_UP；本表用于复核手算透视表、会计重分类、税额和尾差。", 6)
    headers = ["序号", "计算项目", "业务公式", "Excel核对公式", "核对结果", "说明"]
    ws.append(headers)
    gross = Decimal("58616.46")
    rows = [
        [1, "单笔不含税金额", "ROUND_HALF_UP(有效价税合计 ÷ 1.03, 2)", "示例：=ROUND((企业实付金额-企业实退金额)/1.03,2)", "逐笔计算", "先按单笔保留2位"],
        [2, "研发项目金额", "单笔不含税金额 × 项目比例", "示例：=不含税金额*项目比例", "高精度", "订单层不再次舍入，凭证组合汇总后保留2位"],
        [3, "非研发金额", "单笔不含税金额 × 非研发占比", "不含税金额×(1-全部研发项目比例合计)", "高精度", "无姓名加岗位完全一致的工时记录时，非研发占比=100%"],
        [4, "人员与工时匹配", "姓名匹配人员表；会计类别默认取部门或YAML受控覆盖；再以姓名＋会计类别精确匹配工时", "不适用", "确定性规则", "岗位不一致视为无匹配工时，不读取同名其他岗位比例"],
        [5, "月度价税合计", "SUM(有效订单有效价税合计)", "=58616.46", gross, "930笔有效订单"],
        [6, "月度不含税金额", "SUM(逐笔ROUND_HALF_UP(有效价税合计 ÷ 1.03, 2))", "='手算透视复核'!J13", Decimal("56909.24"), "与订单明细AZ列逐笔加总完全一致"],
        [7, "税额", "月度价税合计－逐笔不含税金额合计", "=E8-E9", Decimal("1707.22"), "科目2221.04 待认证进项税额"],
        [8, "费用分录舍入前复核", "手算透视表全部项目及非研发金额合计", "='手算透视复核'!J13", Decimal("56909.24"), "逐笔不含税舍入后再按比例分配"],
        [9, "费用分录舍入后合计", "SUM(各凭证组合ROUND_HALF_UP汇总金额)", "=SUM('凭证分录核对'!F4:F27)", Decimal("56909.26"), "尾差调整前"],
        [10, "尾差", "逐笔不含税金额合计－费用分录舍入后合计", "=ROUND(E9-E12,2)", Decimal("-0.02"), "绝对值小于1元，调整至6602.07/职能008/非研发00009"],
        [11, "费用借方最终合计", "费用分录舍入后合计＋尾差", "=E12+E13", Decimal("56909.24"), "必须等于月度不含税金额"],
        [12, "借方合计", "费用借方最终合计＋税额", "=E14+E10", Decimal("58616.46"), "必须等于贷方"],
        [13, "贷方合计", "应付账款贷方", "=58616.46", Decimal("58616.46"), "科目2202，供应商SC00020"],
    ]
    for row in rows:
        ws.append(row)
    style_table(ws, 3, 1, 6, ws.max_row)
    for row in range(4, ws.max_row + 1):
        ws.cell(row, 5).number_format = "0.00"
    ws.freeze_panes = "A4"
    widths = [8, 22, 48, 42, 16, 48]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # 手算透视复核
    ws = wb.create_sheet("手算透视复核")
    add_title(ws, "手算透视复核", "数据取自滴滴7月网约车.xlsx的手算透视区域；金额保留原始高精度，用于凭证组合汇总。", 10)
    headers = ["管销研", "Coding", "Cowork", "AI4SCI", "Multimodel", "Prediction", "Trading", "Evaluation", "非研发", "行合计"]
    ws.append(headers)
    category_rows = {}
    for values in categories:
        ws.append([values[0], *[float(x) for x in values[1:]], f"=SUM(B{ws.max_row+1}:I{ws.max_row+1})"])
        category_rows[values[0]] = ws.max_row
    total_row = ws.max_row + 1
    ws.cell(total_row, 1).value = "总计"
    for col in range(2, 11):
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = f"=SUM({letter}4:{letter}{total_row-1})"
    style_table(ws, 3, 1, 10, total_row)
    for cell in ws[total_row]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(bold=True)
    for row in range(4, total_row + 1):
        for col in range(2, 11):
            ws.cell(row, col).number_format = "0.0000"
    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = 16
    for col in range(2, 11):
        ws.column_dimensions[get_column_letter(col)].width = 15

    # 凭证分录核对
    ws = wb.create_sheet("凭证分录核对")
    add_title(ws, "凭证分录核对", "“核对公式”直接引用手算透视复核；程序金额取修正后的月度凭证。", 8)
    ws.append(["科目编码", "部门编码", "部门名称", "项目编码", "项目名称", "核对公式金额", "程序金额", "结果"])
    cr = category_rows
    formulas = [
        ("5301.01.07", "003", "开发", "00001", "Coding", f"=ROUND('手算透视复核'!B{cr['开发']},2)"),
        ("5301.01.07", "003", "开发", "00002", "Prediction", f"=ROUND('手算透视复核'!F{cr['开发']},2)"),
        ("5301.01.07", "003", "开发", "00003", "AI4SCI", f"=ROUND('手算透视复核'!D{cr['开发']},2)"),
        ("5301.01.07", "003", "开发", "00005", "Multimodel", f"=ROUND('手算透视复核'!E{cr['开发']},2)"),
        ("5301.01.07", "003", "开发", "00006", "Cowork", f"=ROUND('手算透视复核'!C{cr['开发']},2)"),
        ("5301.01.07", "003", "开发", "00008", "Evaluation", f"=ROUND('手算透视复核'!H{cr['开发']},2)"),
        ("5301.01.07", "005", "算法", "00001", "Coding", f"=ROUND('手算透视复核'!B{cr['算法']}+'手算透视复核'!B{cr['研究员']},2)"),
        ("5301.01.07", "005", "算法", "00002", "Prediction", f"=ROUND('手算透视复核'!F{cr['算法']}+'手算透视复核'!F{cr['研究员']},2)"),
        ("5301.01.07", "005", "算法", "00003", "AI4SCI", f"=ROUND('手算透视复核'!D{cr['算法']}+'手算透视复核'!D{cr['研究员']},2)"),
        ("5301.01.07", "005", "算法", "00005", "Multimodel", f"=ROUND('手算透视复核'!E{cr['算法']}+'手算透视复核'!E{cr['研究员']},2)"),
        ("5301.01.07", "005", "算法", "00006", "Cowork", f"=ROUND('手算透视复核'!C{cr['算法']}+'手算透视复核'!C{cr['研究员']},2)"),
        ("5301.01.07", "005", "算法", "00007", "Trading", f"=ROUND('手算透视复核'!G{cr['算法']}+'手算透视复核'!G{cr['研究员']},2)"),
    ]
    support_rows = [cr[x] for x in ["CEO", "产品运营", "商务", "设计", "职能", "质检"]]
    for code, project, col in [("00002", "Prediction", "F"), ("00003", "AI4SCI", "D"), ("00005", "Multimodel", "E"), ("00006", "Cowork", "C"), ("00007", "Trading", "G"), ("00008", "Evaluation", "H")]:
        expr = "+".join(f"'手算透视复核'!{col}{row}" for row in support_rows)
        formulas.append(("5301.01.07", "012", "研发支持部", code, project, f"=ROUND({expr},2)"))
    formulas.extend([
        ("6601.07", "004", "商务", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['商务']},2)"),
        ("6602.07", "001", "CEO", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['CEO']},2)"),
        ("6602.07", "002", "产品运营", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['产品运营']}+'手算透视复核'!I{cr['设计']},2)"),
        ("6602.07", "007", "质检", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['质检']},2)"),
        ("6602.07", "008", "职能", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['职能']},2)+'公式说明'!E13"),
        ("6602.07", "011", "技术管理部", "00009", "非研发", f"=ROUND('手算透视复核'!I{cr['开发']}+'手算透视复核'!I{cr['算法']}+'手算透视复核'!I{cr['研究员']},2)"),
    ])
    for account, dep, dep_name, project, project_name, formula in formulas:
        actual = new[(account, dep, project)]
        ws.append([account, dep, dep_name, project, project_name, formula, float(actual), f'=IF(ABS(F{ws.max_row+1}-G{ws.max_row+1})<0.005,"一致","不一致")'])
    style_table(ws, 3, 1, 8, ws.max_row)
    for row in range(4, ws.max_row + 1):
        ws.cell(row, 6).number_format = "0.00"
        ws.cell(row, 7).number_format = "0.00"
    ws.freeze_panes = "A4"
    for col, width in enumerate([16, 12, 16, 12, 16, 18, 16, 12], 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    # 差异对比
    ws = wb.create_sheet("差异对比")
    add_title(ws, "原凭证与修正凭证差异", "旧金额取月度凭证-2.xlsx；修正金额取最新样例凭证。", 9)
    ws.append(["科目", "部门", "项目", "旧金额", "修正金额", "差额", "核对公式", "原因", "是否为截图标红"])
    checks = [
        (("5301.01.07", "005", "00002"), "算法", "Prediction", "算法Prediction＋研究员Prediction", "研究员研发金额应归入算法005", "否"),
        (("5301.01.07", "012", "00002"), "研发支持部", "Prediction", "产品运营等支持岗Prediction合计", "姓名加岗位精确匹配后移除错误归入金额", "是"),
        (("5301.01.07", "012", "00005"), "研发支持部", "Multimodel", "产品运营Multimodel", "质检与质检员岗位不一致，按非研发100%", "是"),
        (("6602.07", "002", "00009"), "产品运营", "非研发", "产品运营非研发＋设计非研发", "补入手算岗位为产品运营和设计的非研发金额", "是"),
        (("6602.07", "007", "00009"), "质检", "非研发", "质检全部非研发", "姓名加岗位未完全匹配，质检按非研发100%", "是"),
        (("6602.07", "008", "00009"), "职能", "非研发", "职能非研发＋尾差-0.02", "移出应归算法、产品运营人员，并加入尾差", "是"),
        (("6602.07", "011", "00009"), "技术管理部", "非研发", "开发＋算法＋研究员非研发", "补入算法、研究员岗位未匹配工时的非研发金额", "是"),
    ]
    for key, dep_name, project_name, formula_text, reason, red in checks:
        old_value = old.get(key, Decimal("0"))
        new_value = new.get(key, Decimal("0"))
        ws.append([key[0], f"{dep_name}{key[1]}", f"{project_name}{key[2]}", float(old_value), float(new_value), f"=E{ws.max_row+1}-D{ws.max_row+1}", formula_text, reason, red])
    style_table(ws, 3, 1, 9, ws.max_row)
    for row in range(4, ws.max_row + 1):
        for col in (4, 5, 6):
            ws.cell(row, col).number_format = "0.00"
        if ws.cell(row, 9).value == "是":
            for col in range(1, 10):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=LIGHT_RED)
    ws.freeze_panes = "A4"
    for col, width in enumerate([16, 20, 22, 14, 14, 14, 34, 46, 16], 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                font = copy(cell.font)
                font.name = "Arial Unicode MS"
                cell.font = font
        sheet.sheet_view.showGridLines = False
        sheet.auto_filter.ref = sheet.dimensions if sheet.max_row >= 3 else None
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.sheet_properties.pageSetUpPr.fitToPage = True

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    print(build())
