from __future__ import annotations

import io
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import pytest
from openpyxl import Workbook

from models.schemas import AllocationRecord, DidiOrder, PersonnelRecord
from services.config import load_config


@pytest.fixture
def config():
    return load_config()


def matched_item(position: str, ratios: dict[str, str] | None = None, non_rnd: str = "0"):
    projects = ["Coding", "Prediction", "AI4SCI", "Multimodel", "Cowork", "Trading", "Evaluation"]
    rr = {p: Decimal("0") for p in projects}
    for k, v in (ratios or {}).items():
        rr[k] = Decimal(v)
    order = DidiOrder("O1", "张三", Decimal("103"), Decimal("0"), datetime(2026, 7, 1), "用车订单", 3)
    person = PersonnelRecord("张三", position)
    allocation = AllocationRecord("张三", position, rr, Decimal(non_rnd), "工时分摊-2026.7", 4)
    return [(order, person, allocation)]


def make_template_bytes(hidden: bool = True) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "凭证导入"
    technical = ["book.number", "bizdate", "bookeddate", "vouchertype.number", "vouchertype.name", "billno", "edescription", "account.number", "account.name", "debitlocal", "creditlocal", "currency.number", "currency.name", "assgrp", "assgrp.number", "assgrp.name", "assgrp", "assgrp.number", "assgrp.name"]
    chinese = ["*账簿.编码", "*业务日期", "*记账日期", "*凭证字.编码", "凭证字.名称", "凭证号", "摘要", "*科目.编码", "科目.名称", "借方", "贷方", "*币种.货币代码", "币种.名称", "核算维度", "编码", "名称", "核算维度", "编码", "名称"]
    ws.append(["凭证 # gl_voucher"])
    ws.append(["说明"])
    ws.append(technical)
    ws.append(chinese)
    ws.append(["旧数据"])
    if hidden:
        h = wb.create_sheet("下拉选项")
        h.sheet_state = "hidden"
        h["A1"] = "保留"
    out = io.BytesIO(); wb.save(out); return out.getvalue()
