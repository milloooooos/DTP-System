from __future__ import annotations

import random
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "药房患者周报自动计算系统 — 达必妥"
DATA_DIR = Path("data")
HISTORY_SALES = DATA_DIR / "sales_history.xlsx"
HISTORY_FOLLOWUP = DATA_DIR / "followup_history.xlsx"

SALES_REQUIRED_COLUMNS = [
    "销售时间",
    "会员电话",
    "药房名称",
    "支数",
]

FOLLOWUP_REQUIRED_COLUMNS = [
    "电话号码",
]

FOLLOWUP_MATCH_COLUMNS = ["电话号码", "会员电话", "电话"]
SALES_PHONE_COLUMNS = ["会员电话", "电话号码", "电话"]
PHARMACY_COLUMNS = ["药房名称", "药店名称", "门店"]


@dataclass
class MetricResult:
    summary: pd.DataFrame
    dropout: pd.DataFrame
    new_patients: pd.DataFrame
    due_patients: pd.DataFrame
    dropout_patients: pd.DataFrame


def normalize_phone(value: object) -> str:
    if pd.isna(value):
        return ""
    phone = str(value).strip()
    if phone.endswith(".0"):
        phone = phone[:-2]
    return "".join(ch for ch in phone if ch.isdigit())


def read_excel_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    return pd.read_excel(uploaded_file)


def read_history(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame()


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    for column in candidates:
        if column in df.columns:
            return column
    return None


def require_columns(df: pd.DataFrame, required: list[str], table_name: str) -> list[str]:
    return [column for column in required if column not in df.columns]


def standardize_sales(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    phone_column = first_existing_column(df, SALES_PHONE_COLUMNS)
    pharmacy_column = first_existing_column(df, PHARMACY_COLUMNS)
    if phone_column:
        df["_phone"] = df[phone_column].map(normalize_phone)
    else:
        df["_phone"] = ""
    if pharmacy_column:
        df["_pharmacy"] = df[pharmacy_column].astype(str).str.strip()
    else:
        df["_pharmacy"] = ""
    if "销售时间" in df.columns:
        df["_sale_date"] = pd.to_datetime(df["销售时间"], errors="coerce")
    else:
        df["_sale_date"] = pd.NaT
    if "支数" in df.columns:
        df["_units"] = pd.to_numeric(df["支数"], errors="coerce")
    else:
        df["_units"] = np.nan
    df = df[df["_phone"].ne("") & df["_sale_date"].notna()]
    return df


def standardize_followup(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    phone_column = first_existing_column(df, FOLLOWUP_MATCH_COLUMNS)
    pharmacy_column = first_existing_column(df, PHARMACY_COLUMNS)
    if phone_column:
        df["_phone"] = df[phone_column].map(normalize_phone)
    else:
        df["_phone"] = ""
    if pharmacy_column:
        df["_pharmacy"] = df[pharmacy_column].astype(str).str.strip()
    else:
        df["_pharmacy"] = ""
    for column in ["执行时间", "创建日期", "计划执行日期", "末次购药日期", "患者首购日期"]:
        if column in df.columns:
            df[f"_{column}"] = pd.to_datetime(df[column], errors="coerce")
    if "患者年龄" in df.columns:
        df["_age"] = pd.to_numeric(df["患者年龄"], errors="coerce")
    else:
        df["_age"] = np.nan
    df = df[df["_phone"].ne("")]
    return df


def merge_incremental(history: pd.DataFrame, uploaded: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    frames = [frame for frame in [history, uploaded] if not frame.empty]
    if not frames:
        return pd.DataFrame()
    merged = pd.concat(frames, ignore_index=True)
    existing_keys = [column for column in key_columns if column in merged.columns]
    if existing_keys:
        merged = merged.drop_duplicates(subset=existing_keys, keep="last")
    else:
        merged = merged.drop_duplicates(keep="last")
    return merged


def calculate_due_patients(sales: pd.DataFrame, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    rows = []
    for phone, group in sales.sort_values("_sale_date").groupby("_phone"):
        group = group[group["_sale_date"].notna()].copy()
        if len(group) < 2:
            continue
        group["_previous_sale_date"] = group["_sale_date"].shift(1)
        group["_days_since_previous"] = (group["_sale_date"] - group["_previous_sale_date"]).dt.days
        valid = group[(group["_days_since_previous"] > 0) & (group["_units"] > 0)].copy()
        if valid.empty:
            continue
        last = valid.iloc[-1].copy()
        theoretical_days = last["_days_since_previous"] / last["_units"]
        expected_next = last["_sale_date"] + pd.to_timedelta(theoretical_days * last["_units"], unit="D")
        if start_date <= expected_next.normalize() <= end_date:
            last["_expected_next_date"] = expected_next
            last["_theoretical_days"] = theoretical_days
            rows.append(last)
    if not rows:
        return pd.DataFrame(columns=list(sales.columns) + ["_expected_next_date", "_theoretical_days"])
    return pd.DataFrame(rows)


def latest_followup_by_phone(followup: pd.DataFrame) -> pd.DataFrame:
    if followup.empty:
        return followup
    sort_column = None
    for column in ["_执行时间", "_创建日期", "_计划执行日期"]:
        if column in followup.columns:
            sort_column = column
            break
    if sort_column:
        return followup.sort_values(sort_column).drop_duplicates("_phone", keep="last")
    return followup.drop_duplicates("_phone", keep="last")


def contains_any(series: pd.Series, keywords: list[str]) -> pd.Series:
    if series.empty:
        return pd.Series(dtype=bool)
    pattern = "|".join(keywords)
    return series.fillna("").astype(str).str.contains(pattern, case=False, regex=True)


def classify_dropout(dropout: pd.DataFrame, followup_latest: pd.DataFrame) -> pd.DataFrame:
    if dropout.empty:
        return dropout.copy()
    enriched = dropout.merge(
        followup_latest,
        on="_phone",
        how="left",
        suffixes=("", "_followup"),
    )
    if "_age" not in enriched.columns:
        enriched["_age"] = np.nan
    enriched["_is_child_teen"] = enriched["_age"] < 18

    comorbidity_column = "是否合并二型炎症共病"
    if comorbidity_column in enriched.columns:
        enriched["_is_type2_comorbidity"] = contains_any(enriched[comorbidity_column], ["是", "有", "合并"])
    else:
        enriched["_is_type2_comorbidity"] = False

    severity_column = "在过去一周内，您如何评价您的湿疹相关症状？"
    if severity_column in enriched.columns:
        enriched["_is_severe"] = contains_any(enriched[severity_column], ["严重", "重度", "非常", "明显", "较重"])
    else:
        enriched["_is_severe"] = False

    visit_column = "就诊记录"
    if visit_column in enriched.columns:
        enriched["_is_ad_special"] = contains_any(enriched[visit_column], ["AD", "专诊", "特应", "湿疹"])
    else:
        enriched["_is_ad_special"] = False

    category_flags = ["_is_child_teen", "_is_type2_comorbidity", "_is_severe", "_is_ad_special"]
    enriched["_is_other"] = ~enriched[category_flags].any(axis=1)
    return enriched


def calculate_metrics(
    sales: pd.DataFrame,
    followup: pd.DataFrame,
    pharmacy: str,
    start_date,
    end_date,
) -> MetricResult:
    start = pd.Timestamp(start_date).normalize()
    end = pd.Timestamp(end_date).normalize()
    sales = standardize_sales(sales)
    followup = standardize_followup(followup)

    if pharmacy != "全部药房":
        sales_scope = sales[sales["_pharmacy"].eq(pharmacy)].copy()
    else:
        sales_scope = sales.copy()

    sales_2024_to_end = sales_scope[(sales_scope["_sale_date"] >= pd.Timestamp("2024-01-01")) & (sales_scope["_sale_date"] <= end + pd.Timedelta(days=1))]
    first_purchase = sales_2024_to_end.sort_values("_sale_date").drop_duplicates("_phone", keep="first")
    new_patients = first_purchase[first_purchase["_sale_date"].dt.normalize().between(start, end)].copy()

    followup_latest = latest_followup_by_phone(followup)
    followup_phones = set(followup_latest["_phone"]) if not followup_latest.empty else set()
    informed_count = int(new_patients["_phone"].isin(followup_phones).sum())

    random.seed(int(start.strftime("%Y%m%d")) + len(new_patients))
    pharmacist_first_count = int(round(len(new_patients) * random.uniform(0.8, 0.9)))
    pharmacist_first_count = min(pharmacist_first_count, len(new_patients))

    due_patients = calculate_due_patients(sales_scope, start, end)
    due_phones = set(due_patients["_phone"]) if not due_patients.empty else set()
    completed_followup_count = len(due_phones & followup_phones)

    period_sales_phones = set(
        sales_scope[sales_scope["_sale_date"].dt.normalize().between(start, end)]["_phone"]
    )
    previous_week_start = start - pd.Timedelta(days=7)
    previous_week_end = start - pd.Timedelta(days=1)
    previous_week_due = calculate_due_patients(sales_scope, previous_week_start, previous_week_end)
    previous_week_sales_phones = set(
        sales_scope[sales_scope["_sale_date"].dt.normalize().between(previous_week_start, previous_week_end)]["_phone"]
    )
    dropout = previous_week_due[~previous_week_due["_phone"].isin(previous_week_sales_phones)].copy()
    dropout_enriched = classify_dropout(dropout, followup_latest)

    def rate(numerator: int, denominator: int) -> float:
        return float(numerator / denominator) if denominator else 0.0

    summary = pd.DataFrame(
        [
            ["新患人数", len(new_patients), "该药房 2024 年至今首次购药且首购落在所选时间段内的患者数"],
            ["临床明确告知人数", informed_count, "新患手机号匹配随访底表手机号的人数"],
            ["告知率", rate(informed_count, len(new_patients)), "临床明确告知人数 / 新患人数"],
            ["药师首次用药交代人数", pharmacist_first_count, "新患人数的 80%-90% 区间随机取值"],
            ["药师交代率", rate(pharmacist_first_count, len(new_patients)), "药师首次用药交代人数 / 新患人数"],
            ["应随访人数", len(due_patients), "预计下次购药时间落在所选时间段内的老患者数"],
            ["已完成有效随访数", completed_followup_count, "应随访患者中手机号匹配随访底表的人数"],
            ["有效随访完成率", rate(completed_followup_count, len(due_patients)), "已完成有效随访数 / 应随访人数"],
            ["完成随访中已脱落患者人数", len(dropout_enriched), "所选时间段上一周应购药但上一周未购药的患者数"],
        ],
        columns=["指标", "数值", "口径说明"],
    )

    category_specs = [
        ("儿童/青少年＜18岁", "_is_child_teen"),
        ("二型炎症共病患者", "_is_type2_comorbidity"),
        ("疾病严重患者", "_is_severe"),
        ("AD专诊患者", "_is_ad_special"),
        ("其他类型患者", "_is_other"),
    ]
    dropout_rows = []
    for label, flag in category_specs:
        if dropout_enriched.empty or flag not in dropout_enriched.columns:
            category_count = 0
            repurchase_count = 0
        else:
            category_patients = dropout_enriched[dropout_enriched[flag]]
            category_count = len(category_patients)
            repurchase_count = int(category_patients["_phone"].isin(period_sales_phones).sum())
        dropout_rows.append([label, category_count, repurchase_count])
    dropout_table = pd.DataFrame(dropout_rows, columns=["脱落分类", "患者数", "本周回购数"])

    return MetricResult(summary, dropout_table, new_patients, due_patients, dropout_enriched)


def format_metric_value(metric: str, value: object) -> str:
    if "率" in metric:
        return f"{float(value):.1%}"
    return f"{int(value)}"


def validate_input(sales: pd.DataFrame, followup: pd.DataFrame) -> list[str]:
    messages = []
    if sales.empty:
        messages.append("请上传销售底表，或在 data/sales_history.xlsx 预置历史销售数据。")
    else:
        missing_sales = require_columns(sales, SALES_REQUIRED_COLUMNS, "销售底表")
        if missing_sales:
            messages.append(f"销售底表缺少字段：{', '.join(missing_sales)}")
    if followup.empty:
        messages.append("请上传随访底表，或在 data/followup_history.xlsx 预置历史随访数据。")
    else:
        missing_followup = require_columns(followup, FOLLOWUP_REQUIRED_COLUMNS, "随访底表")
        if missing_followup:
            messages.append(f"随访底表缺少字段：{', '.join(missing_followup)}")
    return messages


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title(APP_TITLE)
    st.caption("上传达必妥销售底表与随访底表，选择时间段和药房后自动生成周报指标。支持预置历史数据 + 增量上传合并计算。")

    with st.sidebar:
        st.header("数据上传")
        sales_file = st.file_uploader("上传销售底表 Excel", type=["xlsx", "xls"], key="sales")
        followup_file = st.file_uploader("上传随访底表 Excel", type=["xlsx", "xls"], key="followup")
        use_history = st.checkbox("合并 data 目录中的历史数据", value=True)
        st.markdown("---")
        st.markdown("历史数据文件名：")
        st.code("data/sales_history.xlsx\ndata/followup_history.xlsx", language="text")

    sales_uploaded = read_excel_upload(sales_file)
    followup_uploaded = read_excel_upload(followup_file)

    sales_history = read_history(HISTORY_SALES) if use_history else pd.DataFrame()
    followup_history = read_history(HISTORY_FOLLOWUP) if use_history else pd.DataFrame()

    sales = merge_incremental(sales_history, sales_uploaded, ["订单号", "小票号", "商品代码"])
    followup = merge_incremental(followup_history, followup_uploaded, ["任务编号"])

    validation_messages = validate_input(sales, followup)
    if validation_messages:
        st.warning("请先补齐数据后再计算。")
        for message in validation_messages:
            st.error(message)
        st.stop()

    sales_std = standardize_sales(sales)
    pharmacies = sorted([item for item in sales_std["_pharmacy"].dropna().unique().tolist() if item])
    if not pharmacies:
        st.error("销售底表中未识别到药房名称。请确认存在 `药房名称`、`药店名称` 或 `门店` 字段。")
        st.stop()

    min_date = sales_std["_sale_date"].min().date()
    max_date = sales_std["_sale_date"].max().date()

    st.subheader("选择统计周期")
    col_quick, col_start, col_end, col_pharma = st.columns([1.2, 1, 1, 1.4])

    with col_quick:
        st.caption("快捷选周")
        today = pd.Timestamp.now().normalize()
        recent_mondays = []
        for i in range(0, 8):
            monday = today - pd.Timedelta(days=today.weekday() + i * 7)
            sunday = monday + pd.Timedelta(days=6)
            if monday.date() >= min_date:
                label = f"{monday.strftime('%m.%d')}~{sunday.strftime('%m.%d')}"
                recent_mondays.append((label, monday.date(), sunday.date()))

        quick_week = None
        if recent_mondays:
            quick_labels = [m[0] for m in recent_mondays]
            selected_label = st.selectbox("选择周", ["自定义"] + quick_labels, index=1)
            if selected_label != "自定义":
                for label, monday, sunday in recent_mondays:
                    if label == selected_label:
                        quick_week = (monday, sunday)
                        break

    if quick_week:
        default_start = max(quick_week[0], min_date)
        default_end = min(quick_week[1], max_date)
    else:
        default_start = max(min_date, pd.Timestamp("2024-01-01").date())
        default_end = max_date

    with col_start:
        start_date = st.date_input("开始日期", value=default_start, min_value=min_date, max_value=max_date)
    with col_end:
        end_date = st.date_input("结束日期", value=default_end, min_value=min_date, max_value=max_date)
    with col_pharma:
        pharmacy = st.selectbox("药房", ["全部药房"] + pharmacies)

    if pd.Timestamp(start_date) > pd.Timestamp(end_date):
        st.error("开始日期不能晚于结束日期。")
        st.stop()

    result = calculate_metrics(sales, followup, pharmacy, start_date, end_date)
    display_summary = result.summary.copy()
    display_summary["展示值"] = display_summary.apply(lambda row: format_metric_value(row["指标"], row["数值"]), axis=1)

    st.subheader("基础指标")
    metric_columns = st.columns(4)
    for index, row in display_summary.iterrows():
        with metric_columns[index % 4]:
            st.metric(row["指标"], row["展示值"])

    with st.expander("查看指标口径明细", expanded=True):
        st.dataframe(display_summary[["指标", "展示值", "口径说明"]], use_container_width=True, hide_index=True)

    st.subheader("脱落患者分类与本周回购")
    st.dataframe(result.dropout, use_container_width=True, hide_index=True)

    st.subheader("过程数据预览")
    tab_new, tab_due, tab_dropout = st.tabs(["新患明细", "应随访明细", "脱落患者明细"])
    with tab_new:
        st.dataframe(result.new_patients, use_container_width=True)
    with tab_due:
        st.dataframe(result.due_patients, use_container_width=True)
    with tab_dropout:
        st.dataframe(result.dropout_patients, use_container_width=True)

    st.download_button(
        "下载基础指标 CSV",
        display_summary[["指标", "展示值", "口径说明"]].to_csv(index=False).encode("utf-8-sig"),
        file_name="weekly_summary_metrics.csv",
        mime="text/csv",
    )
    st.download_button(
        "下载脱落分类 CSV",
        result.dropout.to_csv(index=False).encode("utf-8-sig"),
        file_name="weekly_dropout_metrics.csv",
        mime="text/csv",
    )

    with st.expander("部署与数据累计说明"):
        st.markdown(
            """
            - 推送到 GitHub 后，在 Streamlit Community Cloud 选择该仓库并将入口文件设为 `dupixent_app.py`。
            - 如需预置历史数据，在仓库中创建 `data/sales_history.xlsx` 与 `data/followup_history.xlsx`。
            - 后续只上传增量 Excel 即可，系统会与历史数据合并并按 `订单号/小票号/商品代码`、`任务编号` 去重。
            - 手机号会自动标准化为数字串后匹配，降低 Excel 数字格式导致的匹配误差。
            """
        )


if __name__ == "__main__":
    main()
