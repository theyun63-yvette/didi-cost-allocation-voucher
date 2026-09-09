from __future__ import annotations

from datetime import date
import hashlib

import pandas as pd
import streamlit as st

from services.config import load_config
from services.engine import finalize_voucher, process_calculation
from services.exporters import file_digest
from services.utils import for_display

st.set_page_config(page_title="滴滴车费自动分摊与凭证生成工具", layout="wide")
config = load_config()
st.title("滴滴车费自动分摊与凭证生成工具")
st.caption(f"本地MVP · 规则版本 {config['version']} · 金额使用 Decimal 与 ROUND_HALF_UP")


def safe_df(rows):
    df = pd.DataFrame(rows)
    for column in df.columns:
        if df[column].dtype == "object":
            df[column] = df[column].map(lambda value: "" if value is None else str(value))
    return df


def blocking_text(result) -> str:
    grouped = {}
    for issue in result.blocking_issues:
        key = (issue.code, issue.reason)
        grouped.setdefault(key, set())
        if issue.employee_name:
            grouped[key].add(issue.employee_name)
    reasons = []
    for (code, reason), names in sorted(grouped.items()):
        employee_text = f"（员工：{'、'.join(sorted(names))}）" if names else ""
        reasons.append(f"{code}：{reason}{employee_text}")
    return "\n\n".join(f"- {reason}" for reason in reasons)


def calculation_key(target_month, files) -> str:
    payload = "|".join([target_month, *[file_digest(file) for file in files]])
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


st.header("步骤1：原始账单计算与人工核对")
month_date = st.date_input("目标月份", value=date.today().replace(day=1))
target_month = month_date.strftime("%Y-%m")
cols = st.columns(3)
with cols[0]:
    personnel_file = st.file_uploader("人员信息表", type=["xlsx"], key="personnel")
with cols[1]:
    didi_file = st.file_uploader("原始滴滴网约车账单", type=["xlsx"], key="didi")
with cols[2]:
    allocation_file = st.file_uploader("研发工时成本分摊表", type=["xlsx"], key="allocation")

calculation_files = [personnel_file, allocation_file, didi_file]
if not all(calculation_files):
    st.info("请先上传人员信息表、研发工时成本分摊表和原始滴滴账单。此阶段不需要凭证模板。")
    st.stop()

with st.spinner("正在从原始账单计算订单分摊明细和汇总……"):
    result = process_calculation(personnel_file, allocation_file, didi_file, target_month, config)

st.subheader("第一阶段数据检查")
metric_cols = st.columns(7)
metric_cols[0].metric("目标月份", target_month)
metric_cols[1].metric("原始订单", result.metrics["raw_order_count"])
metric_cols[2].metric("有效订单", result.metrics["valid_order_count"])
metric_cols[3].metric("排除订单", result.metrics["excluded_order_count"])
metric_cols[4].metric("有效价税合计", f"¥{result.metrics['gross_amount']:.2f}")
metric_cols[5].metric("不含税金额合计", f"¥{result.metrics['monthly_net_amount']:.2f}")
metric_cols[6].metric("阻断异常", result.metrics["blocking_issue_count"])

st.write("识别结果：", result.recognized_sheets)
if result.blocking_issues:
    st.error("第一阶段计算存在阻断性异常，不能提交人工确认：\n\n" + blocking_text(result))
else:
    st.success("原始订单计算、单笔分摊平衡和岗位项目汇总校验均已通过。请下载文件进行人工核对。")

checks = {
    "未匹配人员": sorted({x.employee_name for x in result.issues if x.code in {"E003", "E004", "E005"} and x.employee_name}),
    "重复姓名加岗位": sorted({x.employee_name for x in result.issues if x.code == "E006" and x.employee_name}),
    "工时比例不等于1": sorted({x.employee_name for x in result.issues if x.code == "E008" and x.employee_name}),
    "非研发占比空白按0": sorted({x.employee_name for x in result.issues if x.code == "W007" and x.employee_name}),
    "尚未生效项目缺列按0": sorted({str(x.original_value) for x in result.issues if x.code == "W010"}),
    "同月部分退款按净额计算": result.metrics.get("partial_refund_order_count", 0),
    "同月全额退款排除": result.metrics.get("full_refund_order_count", 0),
    "阻断性退款、负数或红字订单": sum(x.code == "E013" and x.blocking for x in result.issues),
    "重复订单号": sum(x.code == "E011" for x in result.issues),
    "停用部门006": sum(x.code == "E010" for x in result.issues),
}
st.json(checks)

stage1_tabs = st.tabs(["订单分摊明细", "人员和工时匹配", "岗位项目分摊汇总", "清洗日志", "异常清单", "计算平衡检查"])
with stage1_tabs[0]:
    st.dataframe(safe_df([for_display(row) for row in result.order_detail_rows]), width="stretch")
with stage1_tabs[1]:
    st.dataframe(safe_df(result.match_rows), width="stretch")
with stage1_tabs[2]:
    st.dataframe(safe_df([for_display(row) for row in result.summary_rows]), width="stretch")
with stage1_tabs[3]:
    st.dataframe(safe_df([for_display(row) for row in result.cleaning_logs]), width="stretch")
with stage1_tabs[4]:
    st.dataframe(safe_df([for_display(issue.as_dict()) for issue in result.issues]), width="stretch")
with stage1_tabs[5]:
    calculation_balance = [
        {"检查项": "订单分摊明细合计=有效订单不含税金额", "左值": result.metrics["order_detail_allocation_total"], "右值": result.metrics["monthly_net_amount"], "通过": result.metrics["order_detail_allocation_total"] == result.metrics["monthly_net_amount"]},
        {"检查项": "岗位项目汇总合计=有效订单不含税金额", "左值": result.metrics["allocation_summary_total"], "右值": result.metrics["monthly_net_amount"], "通过": result.metrics["allocation_summary_total"] == result.metrics["monthly_net_amount"]},
    ]
    st.dataframe(safe_df([for_display(row) for row in calculation_balance]), width="stretch")

st.subheader("第一阶段文件下载")
for name in ["处理明细.xlsx", "分摊汇总.xlsx", "异常清单.xlsx"]:
    st.download_button(f"下载 {name}", result.output_files[name], file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"stage1-{name}")

run_key = calculation_key(target_month, calculation_files)
confirmed = st.session_state.get("confirmed_calculation_key") == run_key
if st.button("我已人工确认分摊明细和汇总无误，进入凭证计算", type="primary", disabled=bool(result.blocking_issues)):
    st.session_state["confirmed_calculation_key"] = run_key
    st.rerun()

if not confirmed:
    st.info("系统正在等待人工确认。确认前不会计算税额、尾差、会计分录或生成月度凭证。")
    st.stop()

st.success("当前分摊结果已人工确认。输入文件或目标月份发生变化后，需要重新确认。")
st.header("步骤2：会计计算与月度凭证生成")
template_file = st.file_uploader("凭证标准模板", type=["xlsx"], key="template")
if not template_file:
    st.info("请上传凭证标准模板，系统才会执行税额、尾差、会计重分类、借贷平衡和凭证生成。")
    st.stop()

with st.spinner("正在执行已确认数据的会计计算并生成凭证……"):
    result = finalize_voucher(result, template_file, config, confirmed=True)

if result.blocking_issues:
    st.error("第二阶段存在阻断性异常，不能生成最终凭证：\n\n" + blocking_text(result))
else:
    st.success("会计计算和借贷平衡校验通过，已生成最终月度凭证。")

amount_cols = st.columns(5)
amount_cols[0].metric("价税合计", f"¥{result.metrics['gross_amount']:.2f}")
amount_cols[1].metric("月度不含税金额", f"¥{result.metrics['monthly_net_amount']:.2f}")
amount_cols[2].metric("税额", f"¥{result.metrics['tax_amount']:.2f}")
amount_cols[3].metric("凭证舍入尾差", f"¥{result.metrics['tail_difference']:.2f}")
amount_cols[4].metric("借贷合计", f"¥{result.metrics['debit_total']:.2f}")

stage2_tabs = st.tabs(["凭证分录预览", "借贷平衡检查", "第二阶段异常", "运行日志"])
with stage2_tabs[0]:
    st.dataframe(safe_df([for_display(line.as_dict()) for line in result.voucher_lines]), width="stretch")
    tail_logs = [row for row in result.audit_log if row.get("事件") == "尾差调整"]
    if tail_logs:
        st.warning("本次发生尾差调整")
        st.dataframe(safe_df([for_display(row) for row in tail_logs]), width="stretch")
with stage2_tabs[1]:
    voucher_balance = [
        {"检查项": "费用借方合计=月度不含税金额", "左值": result.metrics["expense_after_tail"], "右值": result.metrics["monthly_net_amount"], "通过": result.metrics["expense_after_tail"] == result.metrics["monthly_net_amount"]},
        {"检查项": "费用借方+税额=价税合计", "左值": result.metrics["expense_after_tail"] + result.metrics["tax_amount"], "右值": result.metrics["gross_amount"], "通过": result.metrics["expense_after_tail"] + result.metrics["tax_amount"] == result.metrics["gross_amount"]},
        {"检查项": "借方合计=贷方合计", "左值": result.metrics["debit_total"], "右值": result.metrics["credit_total"], "通过": result.metrics["debit_total"] == result.metrics["credit_total"]},
    ]
    st.dataframe(safe_df([for_display(row) for row in voucher_balance]), width="stretch")
with stage2_tabs[2]:
    st.dataframe(safe_df([for_display(issue.as_dict()) for issue in result.issues]), width="stretch")
with stage2_tabs[3]:
    st.dataframe(safe_df([for_display(row) for row in result.audit_log]), width="stretch")

st.subheader("最终凭证下载")
if result.can_generate_voucher and "月度凭证.xlsx" in result.output_files:
    st.download_button("下载 月度凭证.xlsx", result.output_files["月度凭证.xlsx"], file_name="月度凭证.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
else:
    st.button("生成最终凭证（存在阻断性异常）", disabled=True)
