from decimal import Decimal
import pytest
from services.allocation import allocate_orders
from conftest import matched_item


@pytest.mark.parametrize("position,ratios,non_rnd,expected", [
    ("产品运营", {"Coding":"0.2"}, "0.8", {("5301.01.07","012","00001"), ("6602.07","002","00009")}),
    ("商务", {}, "1", {("6601.07","004","00009")}),
    ("开发", {"Coding":"0.8"}, "0.2", {("5301.01.07","003","00001"), ("6602.07","011","00009")}),
    ("算法", {"Coding":"0.8"}, "0.2", {("5301.01.07","005","00001"), ("6602.07","011","00009")}),
    ("研究员", {"Evaluation":"1"}, "0", {("5301.01.07","005","00008")}),
    ("CEO", {"Prediction":"0.3"}, "0.7", {("5301.01.07","012","00002"), ("6602.07","001","00009")}),
    ("设计岗", {}, "1", {("6602.07","002","00009")}),
    ("质检", {"AI4SCI":"0.1"}, "0.9", {("5301.01.07","012","00003"), ("6602.07","007","00009")}),
    ("职能", {"AI4SCI":"0.1"}, "0.9", {("5301.01.07","012","00003"), ("6602.07","008","00009")}),
])
def test_role_accounting_mappings(config, position, ratios, non_rnd, expected):
    details, issues = allocate_orders(matched_item(position, ratios, non_rnd), config)
    assert not issues
    assert {(x.account_code, x.department_code, x.project_code) for x in details} == expected
