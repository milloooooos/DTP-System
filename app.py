import io
import re
from datetime import date

import pandas as pd
import streamlit as st


OUTPUT_COLUMNS = [
    "品牌",
    "药房",
    "执行时间",
    "患者姓名",
    "联系方式",
    "适应症",
    "会员最近一次门店购药时间",
    "会员最近一次门店购药盒数",
    "延期用药原因",
    "本月未购药的原因",
    "医生建议用药时长",
    "需要补充的原因",
]

BRANDS = ["泰瑞沙", "利普卓", "英飞凡", "荃科得", "优赫得", "凡舒卓"]


st.set_page_config(
    page_title="DTP 自动化问题清单",
    layout="wide",
)


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def normalize_phone(value) -> str:
    text = clean_text(value)
    digits = re.sub(r"\D", "", text)
    return digits[-11:] if len(digits) >= 11 else digits


def normalize_member(value) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text.zfill(7) if text.isdigit() and len(text) < 7 else text


def normalize_store(value) -> str:
    text = clean_text(value)
    replacements = [
        "国药控股",
        "四川医药股份有限公司",
        "四川专业药房连锁有限公司",
        "有限公司",
        "股份",
        "医药",
        "关爱大药房",
        "大药房",
        "药房",
        "（连锁）",
        "(连锁）",
        "(连锁)",
        " ",
    ]
    for item in replacements:
        text = text.replace(item, "")
    return text


def parse_date(value):
    if pd.isna(value) or clean_text(value) == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def first_existing(df: pd.DataFrame, names: list[str]):
    for name in names:
        if name in df.columns:
            return name
    return None


def read_excel_upload(uploaded_file) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    return {
        str(name): frame.dropna(how="all").copy()
        for name, frame in sheets.items()
        if not frame.dropna(how="all").empty
    }


def pick_sheet(sheets: dict[str, pd.DataFrame], preferred: str) -> pd.DataFrame:
    if preferred in sheets:
        return sheets[preferred]
    for name, frame in sheets.items():
        if preferred in name:
            return frame
    return next(iter(sheets.values()))


def detect_brand(row: pd.Series) -> str:
    for column in ["品牌", "商品名称", "适应症"]:
        text = clean_text(row.get(column, ""))
        for brand in BRANDS:
            if brand in text:
                return brand
    return clean_text(row.get("品牌", ""))


def detect_target_month(follow_df: pd.DataFrame) -> str:
    if {"年份", "月份"}.issubset(follow_df.columns):
        years = pd.to_numeric(follow_df["年份"], errors="coerce").dropna()
        months = pd.to_numeric(follow_df["月份"], errors="coerce").dropna()
        if not years.empty and not months.empty:
            return f"{int(years.max()):04d}-{int(months.max()):02d}"
    return date.today().strftime("%Y-%m")


def add_follow_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["__brand"] = df.apply(detect_brand, axis=1)
    df["__name"] = df.get("患者姓名", "").map(clean_text) if "患者姓名" in df else ""
    df["__phone"] = df.get("患者手机号", "").map(normalize_phone) if "患者手机号" in df else ""
    df["__member"] = df.get("会员号", "").map(normalize_member) if "会员号" in df else ""
    df["__store"] = df.get("门店", "").map(normalize_store) if "门店" in df else ""
    df["__date"] = df.get("执行时间", pd.NaT).map(parse_date) if "执行时间" in df else pd.NaT
    df["__key_phone"] = df["__brand"] + "|" + df["__phone"]
    df["__key_member"] = df["__brand"] + "|" + df["__member"]
    df["__key_name_store"] = df["__brand"] + "|" + df["__name"] + "|" + df["__store"]
    return df


def add_sales_keys(df: pd.DataFrame, target_month: str) -> pd.DataFrame:
    df = df.copy()
    df["__brand"] = df.apply(detect_brand, axis=1)
    name_col = first_existing(df, ["会员姓名", "开票抬头", "患者姓名"])
    phone_col = first_existing(df, ["会员电话", "手机号", "患者手机号"])
    member_col = first_existing(df, ["会员号", "会员ID"])
    store_col = first_existing(df, ["药房名称", "门店"])

    df["__name"] = df[name_col].map(clean_text) if name_col else ""
    df["__phone"] = df[phone_col].map(normalize_phone) if phone_col else ""
    df["__member"] = df[member_col].map(normalize_member) if member_col else ""
    df["__store"] = df[store_col].map(normalize_store) if store_col else ""
    df["__key_phone"] = df["__brand"] + "|" + df["__phone"]
    df["__key_member"] = df["__brand"] + "|" + df["__member"]
    df["__key_name_store"] = df["__brand"] + "|" + df["__name"] + "|" + df["__store"]

    if "销售时间" in df.columns:
        df["__sale_date"] = df["销售时间"].map(parse_date)
    else:
        df["__sale_date"] = pd.NaT

    month_columns = [column for column in df.columns if re.fullmatch(r"20\d{2}-\d{2}", str(column))]
    if target_month in month_columns:
        df["__month_qty"] = pd.to_numeric(df[target_month], errors="coerce").fillna(0)
    else:
        df["__month_qty"] = 0

    if "销售数量" in df.columns:
        df["__qty"] = pd.to_numeric(df["销售数量"], errors="coerce").fillna(0)
    else:
        df["__qty"] = df["__month_qty"]

    return df


def build_sales_lookup(sales_df: pd.DataFrame, target_month: str) -> dict[str, dict]:
    lookup: dict[str, dict] = {}
    if sales_df.empty:
        return lookup

    target_period = pd.Period(target_month, freq="M")
    sales_df = sales_df.copy()
    sales_df["__in_month"] = (
        sales_df["__sale_date"].dt.to_period("M").eq(target_period)
        if "__sale_date" in sales_df and sales_df["__sale_date"].notna().any()
        else sales_df["__month_qty"].fillna(0).gt(0)
    )

    sort_col = "__sale_date" if sales_df["__sale_date"].notna().any() else "__month_qty"
    for _, row in sales_df.sort_values(sort_col).iterrows():
        record = {
            "brand": clean_text(row.get("__brand")),
            "name": clean_text(row.get("__name")),
            "store": clean_text(row.get("__store")),
            "phone": clean_text(row.get("__phone")),
            "member": clean_text(row.get("__member")),
            "last_sale_date": row.get("__sale_date", pd.NaT),
            "last_qty": row.get("__qty", ""),
            "purchased_this_month": bool(row.get("__in_month", False)),
        }
        for key_name in ["__key_phone", "__key_member", "__key_name_store"]:
            key = clean_text(row.get(key_name))
            keys = [key]
            if not clean_text(row.get("__brand")) and key.startswith("|"):
                keys.append("ANY" + key)
            for match_key in keys:
                if not match_key or match_key.endswith("|"):
                    continue
                old = lookup.get(match_key)
                if old is None:
                    lookup[match_key] = record.copy()
                else:
                    old["purchased_this_month"] = old["purchased_this_month"] or record["purchased_this_month"]
                    if pd.notna(record["last_sale_date"]):
                        old["last_sale_date"] = record["last_sale_date"]
                        old["last_qty"] = record["last_qty"]
    return lookup


def find_sale(row: pd.Series, lookup: dict[str, dict]) -> dict:
    for key_name in ["__key_phone", "__key_member", "__key_name_store"]:
        key = clean_text(row.get(key_name))
        if key in lookup:
            return lookup[key]
        if "|" in key:
            any_key = "ANY|" + key.split("|", 1)[1]
            if any_key in lookup:
                return lookup[any_key]
    return {}


def yes_value(value) -> bool | None:
    text = clean_text(value)
    if text in {"是", "已购药", "有", "Y", "yes", "Yes"}:
        return True
    if text in {"否", "未购药", "无", "N", "no", "No"}:
        return False
    return None


def purchased_in_target_month(row: pd.Series, target_month: str) -> bool | None:
    follow_purchase = yes_value(row.get("当月是否购药"))
    if follow_purchase is not None:
        return follow_purchase

    last_purchase = parse_date(row.get("会员最近一次门店购药时间"))
    if pd.notna(last_purchase):
        return last_purchase.to_period("M") == pd.Period(target_month, freq="M")
    return None


def build_issue_reason(row: pd.Series, sale: dict, target_month: str) -> str:
    reasons = []
    indication = clean_text(row.get("适应症"))
    delay_reason = clean_text(row.get("患者延迟用药的原因"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    follow_purchase = yes_value(row.get("当月是否购药"))
    inferred_purchase = purchased_in_target_month(row, target_month)
    sale_purchase = sale.get("purchased_this_month")
    purchase_status = inferred_purchase if inferred_purchase is not None else sale_purchase

    if not indication:
        reasons.append("适应症未填写，需要补充适应症")

    if purchase_status is True and no_purchase_reason:
        reasons.append("销售底表显示本月已购药，但随访记录了本月未购药原因，需核实")

    if purchase_status is False and not no_purchase_reason:
        reasons.append("销售底表显示本月未购药，随访未填写本月未购药原因")

    if follow_purchase is False and not no_purchase_reason:
        reasons.append("随访记录为当月未购药，但未填写本月未购药原因")

    if follow_purchase is True and no_purchase_reason:
        reasons.append("随访记录为当月已购药，但仍填写了本月未购药原因，需核实")

    if "推迟" in clean_text(row.get("用药周期状态")) and not delay_reason:
        reasons.append("用药周期状态为推迟购药，但未填写延期用药原因")

    return "；".join(dict.fromkeys(reasons))


def build_issue_table(follow_df: pd.DataFrame, sales_df: pd.DataFrame, target_month: str) -> pd.DataFrame:
    follow_df = add_follow_keys(follow_df)
    sales_df = add_sales_keys(sales_df, target_month) if not sales_df.empty else pd.DataFrame()
    sales_lookup = build_sales_lookup(sales_df, target_month)

    rows = []
    for _, follow_row in follow_df.iterrows():
        if not any(
            [
                clean_text(follow_row.get("患者姓名")),
                clean_text(follow_row.get("患者手机号")),
                clean_text(follow_row.get("会员号")),
            ]
        ):
            continue

        sale = find_sale(follow_row, sales_lookup)
        reason = build_issue_reason(follow_row, sale, target_month)
        if not reason:
            continue

        rows.append(
            {
                "品牌": clean_text(follow_row.get("品牌")) or clean_text(follow_row.get("__brand")) or sale.get("brand", ""),
                "药房": clean_text(follow_row.get("门店")),
                "执行时间": follow_row.get("执行时间", ""),
                "患者姓名": clean_text(follow_row.get("患者姓名")),
                "联系方式": clean_text(follow_row.get("患者手机号")),
                "适应症": clean_text(follow_row.get("适应症")),
                "会员最近一次门店购药时间": follow_row.get("会员最近一次门店购药时间", sale.get("last_sale_date", "")),
                "会员最近一次门店购药盒数": follow_row.get("会员最近一次门店购药盒数", sale.get("last_qty", "")),
                "延期用药原因": clean_text(follow_row.get("患者延迟用药的原因")),
                "本月未购药的原因": clean_text(follow_row.get("本月未购药的原因")),
                "医生建议用药时长": clean_text(follow_row.get("医生建议用药时长")),
                "需要补充的原因": reason,
            }
        )

    result = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if not result.empty:
        result = result.drop_duplicates()
        result = result.sort_values(["品牌", "药房", "患者姓名"], na_position="last")
    return result


def to_excel_bytes(result: pd.DataFrame, follow_df: pd.DataFrame, sales_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="问题表格清单")
        follow_df.head(5000).to_excel(writer, index=False, sheet_name="随访底表预览")
        sales_df.head(5000).to_excel(writer, index=False, sheet_name="销售底表预览")
    return output.getvalue()


st.title("DTP 自动化问题清单")
st.caption("上传销售底表和随访底表，自动合并生成所有品种的问题表格清单。")

with st.sidebar:
    st.header("使用方式")
    st.markdown("1. 上传销售底表")
    st.markdown("2. 上传随访底表")
    st.markdown("3. 选择统计月份")
    st.markdown("4. 下载问题表格清单")

sales_file = st.file_uploader("销售底表", type=["xlsx", "xls"], key="sales")
follow_file = st.file_uploader("随访底表", type=["xlsx", "xls"], key="follow")

if not sales_file or not follow_file:
    st.info("请先上传销售底表和随访底表。")
    st.write("当前支持泰瑞沙、利普卓、英飞凡、荃科得、优赫得、凡舒卓等品种字段。")
    st.stop()

try:
    sales_sheets = read_excel_upload(sales_file)
    follow_sheets = read_excel_upload(follow_file)
    sales_df = pick_sheet(sales_sheets, "销售底表")
    follow_df = pick_sheet(follow_sheets, "随访底表")
except Exception as exc:
    st.error(f"Excel 读取失败：{exc}")
    st.stop()

default_month = detect_target_month(follow_df)
target_month = st.text_input("统计月份", value=default_month, help="格式：YYYY-MM，例如 2026-05")

if not re.fullmatch(r"20\d{2}-\d{2}", target_month):
    st.error("统计月份格式需要是 YYYY-MM，例如 2026-05。")
    st.stop()

result = build_issue_table(follow_df, sales_df, target_month)

metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("随访记录", len(follow_df))
metric2.metric("销售记录", len(sales_df))
metric3.metric("问题数量", len(result))
metric4.metric("覆盖品牌", result["品牌"].nunique() if not result.empty else 0)

st.subheader("所有品种问题表格清单")
if result.empty:
    st.success("没有识别到需要补充的问题。")
else:
    selected_brands = st.multiselect(
        "按品牌筛选",
        options=sorted([brand for brand in result["品牌"].dropna().unique() if brand]),
        default=sorted([brand for brand in result["品牌"].dropna().unique() if brand]),
    )
    view_df = result[result["品牌"].isin(selected_brands)] if selected_brands else result
    st.dataframe(view_df, use_container_width=True, hide_index=True)

    st.download_button(
        label="下载问题表格清单",
        data=to_excel_bytes(view_df, follow_df, sales_df),
        file_name=f"DTP问题表格清单_{target_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
