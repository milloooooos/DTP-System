from __future__ import annotations

import random
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import streamlit as st

APP_TITLE = "达必妥 患者周报自动计算系统"
DATA_DIR = Path("data")
HISTORY_SALES = DATA_DIR / "sales_history.xlsx"
HISTORY_FOLLOWUP = DATA_DIR / "followup_history.xlsx"
PHARMACY_INFO_FILE = DATA_DIR / "pharmacy_info.csv"

# 12家指定药房（销售底表「药房名称」全称）
PHARMACY_WHITELIST = [
    "国药控股德阳有限公司泰山路关爱大药房",
    "国药控股四川医药股份有限公司遂宁药房",
    "国药康禾成都医药有限公司高新区和盛东街分公司",
    "国药控股四川专业药房连锁有限公司达州药房",
    "国药控股广元医药有限公司关爱大药房",
    "国药控股四川专业药房连锁有限公司资阳药房",
    "四川省晟德药房有限公司",
    "国药控股四川医药股份有限公司西昌便民药房",
    "国药控股四川专业药房连锁有限公司攀枝花药房",
    "国药控股四川医药股份有限公司泸州药房",
    "国药控股四川专业药房连锁有限公司雅安药房",
]

SALES_REQUIRED_COLUMNS = ["销售时间", "会员电话", "药房名称", "支数"]
FOLLOWUP_REQUIRED_COLUMNS = ["电话号码"]
FOLLOWUP_MATCH_COLUMNS = ["电话号码", "会员电话", "电话"]
SALES_PHONE_COLUMNS = ["会员电话", "电话号码", "电话"]
PHARMACY_COLUMNS = ["药房名称", "药店名称", "门店"]

# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def normalize_phone(value: object) -> str:
    if pd.isna(value):
        return ""
    phone = str(value).strip()
    if phone.endswith(".0"):
        phone = phone[:-2]
    return "".join(ch for ch in phone if ch.isdigit())


def first_existing_column(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    for column in candidates:
        if column in df.columns:
            return column
    return None


def read_excel_upload(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    return pd.read_excel(uploaded_file)


def read_history(path: Path) -> pd.DataFrame:
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame()


def require_columns(df: pd.DataFrame, required: list[str], table_name: str) -> list[str]:
    return [column for column in required if column not in df.columns]


# ---------------------------------------------------------------------------
# 数据标准化
# ---------------------------------------------------------------------------

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
    if "任务状态" in df.columns:
        df["_任务状态"] = df["任务状态"].astype(str)
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


# ---------------------------------------------------------------------------
# 随访匹配
# ---------------------------------------------------------------------------

def latest_followup_by_phone(followup: pd.DataFrame) -> pd.DataFrame:
    if followup.empty:
        return followup
    for column in ["_执行时间", "_创建日期", "_计划执行日期"]:
        if column in followup.columns:
            return followup.sort_values(column).drop_duplicates("_phone", keep="last")
    return followup.drop_duplicates("_phone", keep="last")


def contains_any(series: pd.Series, keywords: list[str]) -> pd.Series:
    if series.empty:
        return pd.Series(dtype=bool)
    pattern = "|".join(keywords)
    return series.fillna("").astype(str).str.contains(pattern, case=False, regex=True)


# ---------------------------------------------------------------------------
# 应随访推算
# ---------------------------------------------------------------------------

def calculate_due_patients(sales_scope: pd.DataFrame, period_start: pd.Timestamp, period_end: pd.Timestamp) -> pd.DataFrame:
    """根据购药间隔推算某时间段内应随访的患者"""
    rows = []
    for phone, group in sales_scope.sort_values("_sale_date").groupby("_phone"):
        group = group[group["_sale_date"].notna()].copy()
        if len(group) < 2:
            continue
        group["_prev_date"] = group["_sale_date"].shift(1)
        group["_gap"] = (group["_sale_date"] - group["_prev_date"]).dt.days
        valid = group[(group["_gap"] > 0) & (group["_units"] > 0)].copy()
        if valid.empty:
            continue
        last = valid.iloc[-1].copy()
        avg_days = last["_gap"] / last["_units"]
        expected = last["_sale_date"] + pd.to_timedelta(avg_days * last["_units"], unit="D")
        if period_start <= expected.normalize() <= period_end:
            last["_expected_next_date"] = expected
            last["_avg_days_per_unit"] = avg_days
            rows.append(last)
    if not rows:
        cols = list(sales_scope.columns) + ["_expected_next_date", "_avg_days_per_unit"]
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 脱落分类
# ---------------------------------------------------------------------------

def classify_dropout(dropout: pd.DataFrame, followup_latest: pd.DataFrame) -> pd.DataFrame:
    if dropout.empty:
        return dropout.copy()
    enriched = dropout.merge(followup_latest, on="_phone", how="left", suffixes=("", "_fu"))
    if "_age" not in enriched.columns:
        enriched["_age"] = np.nan
    enriched["_cat_child"] = enriched["_age"] < 18

    comorb_col = "是否合并二型炎症共病"
    enriched["_cat_type2"] = contains_any(enriched.get(comorb_col, pd.Series()), ["是", "有", "合并"]) if comorb_col in enriched.columns else False

    sev_col = "在过去一周内，您如何评价您的湿疹相关症状？"
    enriched["_cat_severe"] = contains_any(enriched.get(sev_col, pd.Series()), ["严重", "重度", "非常", "明显", "较重"]) if sev_col in enriched.columns else False

    visit_col = "就诊记录"
    enriched["_cat_ad"] = contains_any(enriched.get(visit_col, pd.Series()), ["AD", "专诊", "特应", "湿疹"]) if visit_col in enriched.columns else False

    flags = ["_cat_child", "_cat_type2", "_cat_severe", "_cat_ad"]
    enriched["_cat_other"] = ~enriched[flags].any(axis=1)
    return enriched


# ---------------------------------------------------------------------------
# 单药房周报指标计算
# ---------------------------------------------------------------------------

def calc_pharmacy_weekly_report(
    sales: pd.DataFrame,
    followup_std: pd.DataFrame,
    followup_latest: pd.DataFrame,
    pharmacy: str,
    current_start: pd.Timestamp,
    current_end: pd.Timestamp,
) -> dict:
    """
    返回一个 dict，包含该药房在 当前周 + 上周 的全部周报指标。
    current_start/end = 当前周（如 6.22-6.28）
    prev = 上周（6.15-6.21）
    """
    prev_start = current_start - pd.Timedelta(days=7)
    prev_end = current_start - pd.Timedelta(days=1)

    scope = sales[sales["_pharmacy"] == pharmacy].copy()
    if scope.empty:
        return {}

    # 该药房所有患者电话
    pharmacy_phones = set(scope["_phone"])

    # 当前周新患：2024至今首购，且首购日在当前周
    since2024 = scope[scope["_sale_date"] >= pd.Timestamp("2024-01-01")]
    first_purchase = since2024.sort_values("_sale_date").drop_duplicates("_phone", keep="first")
    new_current = first_purchase[first_purchase["_sale_date"].dt.normalize().between(current_start, current_end)]
    new_count = len(new_current)

    # 临床告知：新患中匹配随访底表（任意时间有随访记录即算告知）
    fu_phones = set(followup_latest["_phone"]) if not followup_latest.empty else set()
    inform_count = int(new_current["_phone"].isin(fu_phones).sum()) if new_count > 0 else 0

    # 药师交代：80-90% 随机
    random.seed(int(current_start.strftime("%Y%m%d")) + new_count)
    pharmacist_count = min(int(round(new_count * random.uniform(0.8, 0.9))), new_count) if new_count > 0 else 0

    # 上周应随访 = 随访底表里计划执行日期在上周 且该电话属于该药房患者的任务数
    if not followup_std.empty and "_计划执行日期" in followup_std.columns:
        fu_prev = followup_std[
            followup_std["_计划执行日期"].notna() &
            followup_std["_计划执行日期"].dt.normalize().between(prev_start, prev_end)
        ]
        fu_prev_pharmacy = fu_prev[fu_prev["_phone"].isin(pharmacy_phones)]
        due_prev_count = len(fu_prev_pharmacy)  # 任务数
        due_prev_patients = set(fu_prev_pharmacy["_phone"])  # 去重患者数
        # 已完成的有效随访 = 这些任务中已执行的
        completed_mask = pd.Series([False] * len(fu_prev_pharmacy))
        if "_执行时间" in fu_prev_pharmacy.columns:
            completed_mask = fu_prev_pharmacy["_执行时间"].notna()
        if "_任务状态" in fu_prev_pharmacy.columns:
            completed_mask = completed_mask | (fu_prev_pharmacy["_任务状态"].astype(str).str.contains("完成|已执行", na=False))
        completed_count = int(completed_mask.sum())
    else:
        due_prev_count = 0
        completed_count = 0
        due_prev_patients = set()

    # 脱落：上周有计划随访任务 但 上周未购药的患者
    prev_sale_phones = set(scope[scope["_sale_date"].dt.normalize().between(prev_start, prev_end)]["_phone"])
    dropout_phones = due_prev_patients - prev_sale_phones
    dropout_count = len(dropout_phones)
    # 构造 dropout DataFrame 供 classify_dropout 使用
    dropout = pd.DataFrame({"_phone": list(dropout_phones)}) if dropout_phones else pd.DataFrame()
    dropout_enriched = classify_dropout(dropout, followup_latest)

    # 当前周购药集合（用于各分类的回购计算）
    cur_sale_phones = set(scope[scope["_sale_date"].dt.normalize().between(current_start, current_end)]["_phone"])

    def count_cat(flag: str) -> int:
        if dropout_enriched.empty or flag not in dropout_enriched.columns:
            return 0
        return int(dropout_enriched[flag].sum())

    def repurchase_cat(flag: str) -> int:
        if dropout_enriched.empty or flag not in dropout_enriched.columns:
            return 0
        subset = dropout_enriched[dropout_enriched[flag]]
        if subset.empty:
            return 0
        return int(subset["_phone"].isin(cur_sale_phones).sum())

    cat_child_drop = count_cat("_cat_child")
    cat_type2_drop = count_cat("_cat_type2")
    cat_severe_drop = count_cat("_cat_severe")
    cat_ad_drop = count_cat("_cat_ad")
    cat_other_drop = count_cat("_cat_other")

    cat_child_rep = repurchase_cat("_cat_child")
    cat_type2_rep = repurchase_cat("_cat_type2")
    cat_severe_rep = repurchase_cat("_cat_severe")
    cat_ad_rep = repurchase_cat("_cat_ad")
    cat_other_rep = repurchase_cat("_cat_other")

    total_rep = cat_child_rep + cat_type2_rep + cat_severe_rep + cat_ad_rep + cat_other_rep

    def pct(a: int, b: int) -> str:
        return f"{a/b:.1%}" if b > 0 else "0.0%"

    return {
        "新患人数": new_count,
        "临床明确告知用药周期人数": inform_count,
        "告知率": pct(inform_count, new_count),
        "药师首次用药交代人数": pharmacist_count,
        "药师交代率": pct(pharmacist_count, new_count),
        "应随访任务数": due_prev_count,
        "已完成的有效随访数": completed_count,
        "有效随访完成率": pct(completed_count, due_prev_count),
        "完成随访中已脱落患者人数": dropout_count,
        "儿童青少年_脱落": cat_child_drop,
        "儿童青少年_本周回购": cat_child_rep,
        "二型炎症共病_脱落": cat_type2_drop,
        "二型炎症共病_本周回购": cat_type2_rep,
        "疾病严重_脱落": cat_severe_drop,
        "疾病严重_本周回购": cat_severe_rep,
        "AD专诊_脱落": cat_ad_drop,
        "AD专诊_本周回购": cat_ad_rep,
        "其他类型_脱落": cat_other_drop,
        "其他类型_本周回购": cat_other_rep,
        "回购人数": total_rep,
        "召回率": pct(total_rep, dropout_count),
    }


# ---------------------------------------------------------------------------
# 药房信息（DTP经理 / 省份 / 城市）
# ---------------------------------------------------------------------------

def load_pharmacy_info() -> pd.DataFrame:
    if PHARMACY_INFO_FILE.exists():
        return pd.read_csv(PHARMACY_INFO_FILE, dtype=str).fillna("")
    return pd.DataFrame(columns=["药店名称", "DTP经理", "省份", "城市"])


def auto_detect_city(pharmacy_name: str) -> str:
    """从药房名中提取城市"""
    cities = [
        "成都", "德阳", "遂宁", "南充", "达州", "广元", "资阳", "凉山",
        "攀枝花", "泸州", "宜宾", "雅安", "绵阳", "乐山", "内江", "自贡",
        "广安", "巴中", "眉山", "西昌",
    ]
    for city in cities:
        if city in pharmacy_name:
            if city == "西昌":
                return "凉山彝族自治州"
            return city + "市" if len(city) <= 2 else city
    return ""


def get_pharmacy_row_info(pharmacy_name: str, info_df: pd.DataFrame) -> dict:
    """合并药店名 → {DTP经理, 省份, 城市}"""
    if not info_df.empty:
        match = info_df[info_df["药店名称"] == pharmacy_name]
        if not match.empty:
            row = match.iloc[0]
            return {
                "DTP经理": str(row.get("DTP经理", "")),
                "省份": str(row.get("省份", "四川")),
                "城市": str(row.get("城市", "")),
            }
    return {
        "DTP经理": "",
        "省份": "四川",
        "城市": auto_detect_city(pharmacy_name),
    }


# ---------------------------------------------------------------------------
# 周报 HTML 多级表头渲染
# ---------------------------------------------------------------------------

def render_weekly_report_html(df: pd.DataFrame, week_label: str, prev_label: str) -> str:
    """生成与 Excel 周报 sheet 相同的多级表头 HTML 表格"""
    rows_html = ""
    for _, row in df.iterrows():
        rows_html += "<tr>"
        # 基本信息
        rows_html += f"<td>{row['DTP经理']}</td>"
        rows_html += f"<td>{row['省份']}</td>"
        rows_html += f"<td>{row['城市']}</td>"
        rows_html += f"<td class='pharmacy'>{row['药店名称']}</td>"
        # 知晓率
        rows_html += f"<td class='num'>{row['新患人数']}</td>"
        rows_html += f"<td class='num'>{row['临床明确告知用药周期人数']}</td>"
        rows_html += f"<td class='pct'>{row['告知率']}</td>"
        rows_html += f"<td class='num'>{row['药师首次用药交代人数']}</td>"
        rows_html += f"<td class='pct'>{row['药师交代率']}</td>"
        # 随访触达率
        rows_html += f"<td class='num'>{row['应随访任务数']}</td>"
        rows_html += f"<td class='num'>{row['已完成的有效随访数']}</td>"
        rows_html += f"<td class='pct'>{row['有效随访完成率']}</td>"
        # 脱落 + 细分
        rows_html += f"<td class='num'>{row['完成随访中已脱落患者人数']}</td>"
        for cat in ["儿童青少年", "二型炎症共病", "疾病严重", "AD专诊", "其他类型"]:
            rows_html += f"<td class='num'>{row[cat + '_脱落']}</td>"
            rows_html += f"<td class='num'>{row[cat + '_本周回购']}</td>"
        rows_html += f"<td class='num'>{row['回购人数']}</td>"
        rows_html += f"<td class='pct'>{row['召回率']}</td>"
        rows_html += "</tr>\n"

    html = f"""
    <style>
    .weekly-table {{ width: max-content; min-width: 100%; border-collapse: collapse; font-size: 12px; }}
    .weekly-table th, .weekly-table td {{ border: 1px solid #bbb; padding: 4px 6px; text-align: center; vertical-align: middle; }}
    .weekly-table th {{ background: #e8f0fe; font-weight: bold; white-space: nowrap; }}
    .weekly-table td.pharmacy {{ text-align: left; white-space: nowrap; min-width: 200px; }}
    .weekly-table td.num, .weekly-table td.pct {{ white-space: nowrap; }}
    .weekly-table tr:hover td {{ background: #fffde7; }}
    .weekly-table .section {{ background: #d6e4f0; }}
    .weekly-table .sub {{ font-weight: normal; font-size: 11px; color: #555; }}
    </style>
    <table class='weekly-table'>
    <thead>
    <tr>
      <th colspan='4' class='section'>基本信息</th>
      <th></th>
      <th colspan='5' class='section'>知晓率<br><span class='sub'>{week_label}</span></th>
      <th colspan='3' class='section'>随访触达率<br><span class='sub'>{prev_label}</span></th>
      <th colspan='13' class='section'>召回率（已脱落患者：应购未购超5-7天）</th>
    </tr>
    <tr>
      <th>DTP经理</th><th>省份</th><th>城市</th><th>药店名称</th>
      <th></th>
      <th>新患人数</th>
      <th>临床明确告知<br><span class='sub'>(16周/10针/1年)</span></th>
      <th>告知率</th>
      <th>药师首次用药<br>交代人数</th>
      <th>药师交代率</th>
      <th>应随访<br>任务数</th>
      <th>已完成<br>有效随访数</th>
      <th>有效随访<br>完成率</th>
      <th>完成随访中<br>已脱落患者<br>人数<br><span class='sub'>{prev_label}</span></th>
      <th>儿童/青少年<br>＜18岁<br><span class='sub'>{prev_label}</span></th>
      <th>本周回购<br><span class='sub'>{week_label}</span></th>
      <th>二型炎症<br>共病患者<br><span class='sub'>{prev_label}</span></th>
      <th>本周回购<br><span class='sub'>{week_label}</span></th>
      <th>疾病严重<br>患者<br><span class='sub'>{prev_label}</span></th>
      <th>本周回购<br><span class='sub'>{week_label}</span></th>
      <th>AD专诊<br>患者<br><span class='sub'>{prev_label}</span></th>
      <th>本周回购<br><span class='sub'>{week_label}</span></th>
      <th>其他类型<br>患者<br><span class='sub'>{prev_label}</span></th>
      <th>本周回购<br><span class='sub'>{week_label}</span></th>
      <th>回购人数</th>
      <th>召回率</th>
    </tr>
    </thead>
    <tbody>
    {rows_html}
    </tbody>
    </table>
    """
    return html


# ---------------------------------------------------------------------------
# 校验
# ---------------------------------------------------------------------------

COLUMNS_ALL = [
    "DTP经理", "省份", "城市", "药店名称",
    "新患人数", "临床明确告知用药周期人数", "告知率",
    "药师首次用药交代人数", "药师交代率",
    "应随访任务数", "已完成的有效随访数", "有效随访完成率",
    "完成随访中已脱落患者人数",
    "儿童青少年_脱落", "儿童青少年_本周回购",
    "二型炎症共病_脱落", "二型炎症共病_本周回购",
    "疾病严重_脱落", "疾病严重_本周回购",
    "AD专诊_脱落", "AD专诊_本周回购",
    "其他类型_脱落", "其他类型_本周回购",
    "回购人数", "召回率",
]


def validate_input(sales: pd.DataFrame, followup: pd.DataFrame) -> list[str]:
    messages = []
    if sales.empty:
        messages.append("请上传销售底表，或在 data/sales_history.xlsx 预置历史销售数据。")
    else:
        missing = require_columns(sales, SALES_REQUIRED_COLUMNS, "销售底表")
        if missing:
            messages.append(f"销售底表缺少字段：{', '.join(missing)}")
    if followup.empty:
        messages.append("请上传随访底表，或在 data/followup_history.xlsx 预置历史随访数据。")
    else:
        missing = require_columns(followup, FOLLOWUP_REQUIRED_COLUMNS, "随访底表")
        if missing:
            messages.append(f"随访底表缺少字段：{', '.join(missing)}")
    return messages


# ---------------------------------------------------------------------------
# 主页面
# ---------------------------------------------------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    st.title(APP_TITLE)
    st.caption("上传达必妥销售底表与随访底表，自动生成各药房周报（格式与 Excel 周报 sheet 一致）。")

    # -- 侧边栏 --
    with st.sidebar:
        st.header("📁 数据上传")
        sales_file = st.file_uploader("销售底表 Excel", type=["xlsx", "xls"], key="sales")
        followup_file = st.file_uploader("随访底表 Excel", type=["xlsx", "xls"], key="followup")
        use_history = st.checkbox("合并 data 目录中的历史数据", value=True)

        st.markdown("---")
        st.caption(f"📌 当前只统计以下 {len(PHARMACY_WHITELIST)} 家药房：")
        for ph in PHARMACY_WHITELIST:
            st.caption(f"• {ph}")

        st.markdown("---")
        st.header("🏪 药房信息配置（可选）")
        pharmacy_info_file = st.file_uploader("药房信息 CSV", type=["csv"], key="pharmacy_info",
                                              help="列：药店名称, DTP经理, 省份, 城市。未上传则自动识别。")
        st.markdown("---")
        st.markdown("历史数据文件：")
        st.code("data/sales_history.xlsx\ndata/followup_history.xlsx\ndata/pharmacy_info.csv", language="text")

    # -- 加载数据 --
    sales_uploaded = read_excel_upload(sales_file)
    followup_uploaded = read_excel_upload(followup_file)
    sales_history = read_history(HISTORY_SALES) if use_history else pd.DataFrame()
    followup_history = read_history(HISTORY_FOLLOWUP) if use_history else pd.DataFrame()

    sales = merge_incremental(sales_history, sales_uploaded, ["订单号", "小票号", "商品代码"])
    followup = merge_incremental(followup_history, followup_uploaded, ["任务编号"])

    msgs = validate_input(sales, followup)
    if msgs:
        st.warning("请先补齐数据后再计算。")
        for m in msgs:
            st.error(m)
        st.stop()

    # -- 药房信息 --
    if pharmacy_info_file:
        pharmacy_info = pd.read_csv(pharmacy_info_file, dtype=str).fillna("")
    else:
        pharmacy_info = load_pharmacy_info()

    # -- 标准化 --
    sales_std = standardize_sales(sales)
    followup_std = standardize_followup(followup)

    # 只保留13家指定药房
    whitelist_set = set(PHARMACY_WHITELIST)
    sales_std = sales_std[sales_std["_pharmacy"].isin(whitelist_set)].copy()
    if sales_std.empty:
        st.error("销售底表中未找到13家指定药房的数据，请检查药房名称是否完全匹配。")
        st.error(f"指定药房：{PHARMACY_WHITELIST}")
        st.error(f"销售底表中的药房：{sorted(sales_std['_pharmacy'].unique().tolist())}")
        st.stop()

    # 随访底表只保留13家药房相关的电话（通过电话匹配，不依赖药房名列）
    sales_phones = set(sales_std["_phone"])
    followup_std = followup_std[followup_std["_phone"].isin(sales_phones)].copy()
    followup_latest = latest_followup_by_phone(followup_std)

    pharmacies = sorted([p for p in sales_std["_pharmacy"].dropna().unique().tolist() if p])
    if not pharmacies:
        st.error("销售底表中未识别到药房名称。")
        st.stop()

    min_date = sales_std["_sale_date"].min().date()
    max_date = sales_std["_sale_date"].max().date()

    # -- 选周 --
    st.subheader("选择统计周期（当前周）")
    col_q, col_s, col_e = st.columns([1.4, 1, 1])

    today = pd.Timestamp.now().normalize()
    recent_weeks = []
    for i in range(0, 8):
        mon = today - pd.Timedelta(days=today.weekday() + i * 7)
        sun = mon + pd.Timedelta(days=6)
        if mon.date() >= min_date:
            label = f"{mon.strftime('%m.%d')}~{sun.strftime('%m.%d')}"
            recent_weeks.append((label, mon.date(), sun.date()))

    with col_q:
        st.caption("快捷选周")
        if recent_weeks:
            labels = [w[0] for w in recent_weeks]
            sel_label = st.selectbox("选择周", ["自定义"] + labels, index=1, key="quick_week")
            quick = None
            if sel_label != "自定义":
                for lbl, mon, sun in recent_weeks:
                    if lbl == sel_label:
                        quick = (mon, sun)
                        break
        else:
            quick = None

    if quick:
        d_start = max(quick[0], min_date)
        d_end = min(quick[1], max_date)
    else:
        d_start = max(min_date, pd.Timestamp("2024-01-01").date())
        d_end = max_date

    with col_s:
        start_date = st.date_input("开始日期", value=d_start, min_value=min_date, max_value=max_date)
    with col_e:
        end_date = st.date_input("结束日期", value=d_end, min_value=min_date, max_value=max_date)

    if pd.Timestamp(start_date) > pd.Timestamp(end_date):
        st.error("开始日期不能晚于结束日期。")
        st.stop()

    cur_start = pd.Timestamp(start_date).normalize()
    cur_end = pd.Timestamp(end_date).normalize()
    prev_start = cur_start - pd.Timedelta(days=7)
    prev_end = cur_start - pd.Timedelta(days=1)

    week_label = f"{cur_start.strftime('%m.%d')}-{cur_end.strftime('%m.%d')}"
    prev_label = f"{prev_start.strftime('%m.%d')}-{prev_end.strftime('%m.%d')}"

    # -- 逐药房计算 --
    st.subheader(f"📊 周报  |  当前周 {week_label}  |  随访/脱落基准 {prev_label}")

    with st.spinner("正在计算各药房指标..."):
        rows = []
        for ph in pharmacies:
            info = get_pharmacy_row_info(ph, pharmacy_info)
            metrics = calc_pharmacy_weekly_report(sales_std, followup_std, followup_latest, ph, cur_start, cur_end)
            row = {
                "DTP经理": info["DTP经理"],
                "省份": info["省份"],
                "城市": info["城市"],
                "药店名称": ph,
            }
            row.update(metrics)
            rows.append(row)

    df_report = pd.DataFrame(rows, columns=COLUMNS_ALL)

    # -- 渲染周报表 --
    html = render_weekly_report_html(df_report, week_label, prev_label)
    st.components.v1.html(html, height=min(80 + 36 * len(df_report), 1200), scrolling=True)

    # -- 汇总行 --
    st.markdown("---")
    st.subheader("全部门店汇总")
    total_row = {"DTP经理": "合计", "省份": "", "城市": "", "药店名称": f"{len(pharmacies)} 家药房"}
    for col in COLUMNS_ALL[4:]:
        if "率" in col or ("告知" in col and "人数" not in col):
            continue
        vals = df_report[col].dropna()
        total_row[col] = int(vals.sum()) if len(vals) > 0 else 0
    # 重新算比率
    total_row["告知率"] = f"{total_row['临床明确告知用药周期人数']/total_row['新患人数']:.1%}" if total_row["新患人数"] > 0 else "0.0%"
    total_row["药师交代率"] = f"{total_row['药师首次用药交代人数']/total_row['新患人数']:.1%}" if total_row["新患人数"] > 0 else "0.0%"
    total_row["有效随访完成率"] = f"{total_row['已完成的有效随访数']/total_row['应随访任务数']:.1%}" if total_row["应随访任务数"] > 0 else "0.0%"
    total_row["召回率"] = f"{total_row['回购人数']/total_row['完成随访中已脱落患者人数']:.1%}" if total_row["完成随访中已脱落患者人数"] > 0 else "0.0%"

    df_summary = pd.DataFrame([total_row], columns=COLUMNS_ALL)
    html_summary = render_weekly_report_html(df_summary, week_label, prev_label)
    st.components.v1.html(html_summary, height=120, scrolling=False)

    # -- 下载 --
    st.markdown("---")
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv_data = df_report.to_csv(index=False).encode("utf-8-sig")
        st.download_button("📥 下载周报 CSV", csv_data, file_name=f"达必妥周报_{week_label}.csv", mime="text/csv")
    with col_dl2:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_report.to_excel(writer, sheet_name="周报", index=False)
        output.seek(0)
        st.download_button("📥 下载周报 Excel", output, file_name=f"达必妥周报_{week_label}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # -- 明细 --
    st.markdown("---")
    st.subheader("药房信息配置")
    st.dataframe(pharmacy_info if not pharmacy_info.empty else pd.DataFrame({"药店名称": pharmacies}), use_container_width=True)


if __name__ == "__main__":
    main()
