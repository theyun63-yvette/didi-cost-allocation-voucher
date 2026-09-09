from datetime import datetime
from decimal import Decimal
from models.schemas import AllocationDetail, AllocationRecord, DidiOrder, PersonnelRecord, VoucherLine
from services.accounting import build_voucher_lines
from services.matcher import match_orders
from services.validation import validate_voucher


def expense_detail(amount: str):
    return AllocationDetail("O1", "张三", "职能", Decimal("103"), Decimal("100"), "非研发", "非研发", Decimal("1"), Decimal(amount), "6602.07", "日常交通费", "008", "职能", "00009", "CNY", "用车订单", 3)


def order(gross="103"):
    return DidiOrder("O1", "张三", Decimal(gross), Decimal("0"), datetime(2026,7,1), "用车订单", 3)


def test_ratio_not_one_is_blocking(config):
    from parsers.allocation_parser import parse_allocation
    from openpyxl import Workbook
    import io
    wb=Workbook(); ws=wb.active; ws.title="工时分摊-2026.7"
    ws.append(["标题"]); ws.append(["序号","姓名","岗位","Coding","Prediction","AI4SCI","Multimodel","Cowork","Trading","Evaluation","非研发占比"])
    ws.append([1,"张三","开发",0.5,0,0,0,0,0,0,0.4])
    b=io.BytesIO(); wb.save(b)
    result=parse_allocation(b.getvalue(),"2026-07",config)
    assert any(x.code=="E008" and x.blocking for x in result.issues)


def test_duplicate_name_position_blocks_match():
    p=[PersonnelRecord("张三","开发")]
    ratios={"Coding":Decimal("1")}
    a=[AllocationRecord("张三","开发",ratios,Decimal("0"),"s",2),AllocationRecord("张三","开发",ratios,Decimal("0"),"s",3)]
    matched, _, issues=match_orders([order()],p,a)
    assert not matched and any(x.code=="E006" and x.blocking for x in issues)


def test_department_006_is_blocking(config):
    lines=[VoucherLine("费用","6602.07","日常交通费",debit=Decimal("100"),department_code="006",department_name="运营",project_code="00009",project_name="非研发"),VoucherLine("税额","2221.04","待认证进项税额",debit=Decimal("3")),VoucherLine("应付账款","2202","应付账款",credit=Decimal("103"))]
    metrics={"gross_amount":Decimal("103"),"monthly_net_amount":Decimal("100"),"tax_amount":Decimal("3")}
    issues=validate_voucher(lines,metrics,config)
    assert any(x.code=="E010" and x.blocking for x in issues)


def test_tail_099_auto_adjusts(config):
    lines, issues, metrics, audit=build_voucher_lines([expense_detail("99.01")],[order()],config)
    assert not issues
    assert metrics["tail_difference"]==Decimal("0.99")
    assert any(x.tail_adjustment for x in lines)
    assert audit and metrics["expense_after_tail"]==Decimal("100.00")


def test_tail_100_blocks(config):
    lines, issues, metrics, audit=build_voucher_lines([expense_detail("99.00")],[order()],config)
    assert metrics["tail_difference"]==Decimal("1.00")
    assert any(x.code=="E014" and x.blocking for x in issues)


def test_final_voucher_balances(config):
    lines, issues, metrics, _=build_voucher_lines([expense_detail("100")],[order()],config)
    assert not issues
    assert not validate_voucher(lines,metrics,config)
    assert metrics["debit_total"]==metrics["credit_total"]==Decimal("103.00")


def test_no_timesheet_business_is_100_percent_sales(config):
    from services.allocation import allocate_orders
    person = PersonnelRecord("张三", "商务", department_code="004", department_name="商务")
    matched, rows, issues = match_orders([order()], [person], [], project_names=[x for x in config["projects"] if x != config["constants"]["non_rnd_project"]])
    assert not [x for x in issues if x.blocking]
    assert matched[0][2].non_rnd_ratio == Decimal("1")
    details, allocation_issues = allocate_orders(matched, config)
    assert not allocation_issues
    assert {(x.account_code, x.department_code, x.project_code) for x in details} == {("6601.07", "004", "00009")}


def test_no_timesheet_non_business_is_100_percent_management(config):
    from services.allocation import allocate_orders
    person = PersonnelRecord("张三", "产品运营", department_code="002", department_name="产品运营")
    matched, rows, issues = match_orders([order()], [person], [], project_names=[x for x in config["projects"] if x != config["constants"]["non_rnd_project"]])
    assert not [x for x in issues if x.blocking]
    details, allocation_issues = allocate_orders(matched, config)
    assert not allocation_issues
    assert {(x.account_code, x.department_code, x.project_code) for x in details} == {("6602.07", "002", "00009")}


def test_timesheet_requires_exact_name_and_accounting_role(config):
    from services.allocation import allocate_orders
    projects = {name: Decimal("0") for name in config["projects"] if name != config["constants"]["non_rnd_project"]}
    projects["Coding"] = Decimal("0.2")
    person = PersonnelRecord("张三", "产品运营", department_code="002", department_name="产品运营")
    allocation = AllocationRecord("张三", "算法", projects, Decimal("0.8"), "工时分摊-2026.7", 3)
    matched, rows, issues = match_orders([order()], [person], [allocation], project_names=list(projects))
    assert not [x for x in issues if x.blocking]
    assert matched[0][2].non_rnd_ratio == Decimal("1")
    details, allocation_issues = allocate_orders(matched, config)
    assert not allocation_issues
    assert {(x.account_code, x.department_code, x.project_code) for x in details} == {
        ("6602.07", "002", "00009"),
    }


def test_employee_role_override_controls_timesheet_match(config):
    from services.allocation import allocate_orders
    projects = {name: Decimal("0") for name in config["projects"] if name != config["constants"]["non_rnd_project"]}
    projects["Coding"] = Decimal("0.2")
    person = PersonnelRecord("张三", "职能", department_code="008", department_name="职能")
    allocation = AllocationRecord("张三", "算法", projects, Decimal("0.8"), "工时分摊-2026.7", 3)
    matched, _, issues = match_orders(
        [order()], [person], [allocation], project_names=list(projects), employee_role_overrides={"张三": "算法"}
    )
    assert not [x for x in issues if x.blocking]
    details, allocation_issues = allocate_orders(matched, config)
    assert not allocation_issues
    assert {(x.account_code, x.department_code, x.project_code) for x in details} == {
        ("5301.01.07", "005", "00001"),
        ("6602.07", "011", "00009"),
    }


def test_monthly_tax_uses_sum_of_order_level_rounded_net(config):
    details = [expense_detail("199.99")]
    orders = [order("103"), DidiOrder("O2", "李四", Decimal("103.01"), Decimal("0"), datetime(2026,7,2), "用车订单", 4)]
    lines, issues, metrics, _ = build_voucher_lines(details, orders, config)
    assert metrics["gross_amount"] == Decimal("206.01")
    assert metrics["monthly_net_amount"] == Decimal("200.01")
    assert metrics["tax_amount"] == Decimal("6.00")

