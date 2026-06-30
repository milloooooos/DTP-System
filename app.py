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
SPECIAL_FOLLOW_BRANDS = {"优赫得"}
# Brands whose correct list uses RAW (full) store names from follow data "门店"
RAW_STORE_BRANDS = {"泰瑞沙", "利普卓", "英飞凡"}
# Brands whose correct list uses ALIASED (short) store names from sales data "药房名称"
ALIASED_STORE_BRANDS = {"荃科得", "优赫得", "凡舒卓"}
PHARMACY_ALIASES = {
    "国药控股德阳有限公司泰山路关爱大药房": "德阳关爱药房",
    "国药控股四川医药股份有限公司眉山药房": "眉山药房",
    "国药控股巴中医药有限公司兴文关爱大药房": "巴中兴文药房",
    "国药控股内江有限公司第一大药房": "内江第一药房",
    "国药控股广安有限公司广安药房": "广安药房",
    "国药控股四川专业药房连锁有限公司攀枝花药房": "攀枝花药房(连锁）",
    "国药控股四川医药股份有限公司南充药房": "南充药房",
    "国药控股巴中医药有限公司通江店": "巴中通江店",
    "四川省晟德药房有限公司": "成都晟德药房",
    "国药控股巴中医药有限公司平昌关爱店": "巴中平昌店",
    "国药控股德阳有限公司吉康大药房": "德阳吉康药房",
    "国药控股广元医药有限公司关爱大药房": "广元关爱药房",
    "国药控股四川医药股份有限公司西昌便民药房": "西昌药房",
    "国药控股德阳有限公司喜悦大药房": "德阳喜悦药房",
    "国药控股泸州医药有限公司荔城药房": "泸州荔城药房",
    "国药控股昊阳绵阳药业有限公司江油匡山路大药房": "绵阳江油药房",
}
# Scope mapping for 泰瑞沙/利普卓 rolling-window approach: exactly 15 stores from automation 药房简称 sheet
# (excludes 绵阳江油药房 which is NOT in the automation scope)
PHARMACY_SCOPE = {
    "国药控股德阳有限公司泰山路关爱大药房": "德阳关爱药房",
    "国药控股四川医药股份有限公司眉山药房": "眉山药房",
    "国药控股巴中医药有限公司兴文关爱大药房": "巴中兴文药房",
    "国药控股内江有限公司第一大药房": "内江第一药房",
    "国药控股广安有限公司广安药房": "广安药房",
    "国药控股四川专业药房连锁有限公司攀枝花药房": "攀枝花药房(连锁）",
    "国药控股四川医药股份有限公司南充药房": "南充药房",
    "国药控股巴中医药有限公司通江店": "巴中通江店",
    "四川省晟德药房有限公司": "成都晟德药房",
    "国药控股巴中医药有限公司平昌关爱店": "巴中平昌店",
    "国药控股德阳有限公司吉康大药房": "德阳吉康药房",
    "国药控股广元医药有限公司关爱大药房": "广元关爱药房",
    "国药控股四川医药股份有限公司西昌便民药房": "西昌药房",
    "国药控股德阳有限公司喜悦大药房": "德阳喜悦药房",
    "国药控股泸州医药有限公司荔城药房": "泸州荔城药房",
}
# Reverse mapping: aliased → raw, for brands that need raw store names
REVERSE_ALIASES = {v: k for k, v in PHARMACY_ALIASES.items()}


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
    if text in PHARMACY_ALIASES:
        return PHARMACY_ALIASES[text]
    return text.replace(" ", "").replace("（", "(").replace("）", ")")


def parse_date(value):
    if pd.isna(value) or clean_text(value) == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def parse_datetime(value):
    """Parse datetime preserving time component (needed for 优赫得/凡舒卓 expected dates).
    Unlike parse_date which normalizes to midnight, this keeps the original time."""
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


def classify_follow_file(frame: pd.DataFrame, filename: str) -> str:
    if any(brand in filename for brand in SPECIAL_FOLLOW_BRANDS):
        return "优赫得专属表"
    if any(brand in filename for brand in set(BRANDS) - SPECIAL_FOLLOW_BRANDS):
        return "通用随访表"

    columns = set(frame.columns)
    has_doctor_duration = bool({"医生建议用药时长", "医生建议服用时间"} & columns)
    has_month_purchase = "当月是否购药" in columns
    if has_month_purchase and not has_doctor_duration:
        return "优赫得专属表"
    if has_doctor_duration:
        return "通用随访表"
    return "未知随访表"


def combine_follow_files(uploaded_files) -> pd.DataFrame:
    prepared = []
    for uploaded_file in uploaded_files:
        sheets = read_excel_upload(uploaded_file)
        frame = pick_sheet(sheets, "随访底表").copy()
        file_type = classify_follow_file(frame, uploaded_file.name)
        frame["__brand"] = frame.apply(detect_brand, axis=1)
        frame["来源文件"] = uploaded_file.name
        frame["随访表类型"] = file_type
        # Detect if this file is the "报表" (summary) vs "明细报表" (detail)
        # 报表 files have "医生建议用药时长" or "医生建议服用时间" columns (通用随访表)
        has_doctor_duration = bool({"医生建议用药时长", "医生建议服用时间"} & set(frame.columns))
        frame["__is_report"] = has_doctor_duration  # report files are more authoritative
        prepared.append((file_type, frame))
    file_types = {file_type for file_type, _ in prepared}
    should_route_by_brand = len(prepared) > 1 and "优赫得专属表" in file_types

    frames = []
    for file_type, frame in prepared:
        if should_route_by_brand:
            if file_type == "优赫得专属表":
                frame = frame[frame["__brand"].isin(SPECIAL_FOLLOW_BRANDS)].copy()
            else:
                frame = frame[~frame["__brand"].isin(SPECIAL_FOLLOW_BRANDS)].copy()
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    # For brands that appear in both report and detail files (泰瑞沙/利普卓/英飞凡/凡舒卓),
    # prefer the report file's data during dedup by putting report rows first
    # This ensures authoritative data is kept for overlapping patient+store pairs
    return combined


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
    df["__store_raw"] = df.get("门店", "").map(clean_text) if "门店" in df else ""
    df["__date"] = df.get("执行时间", pd.NaT).map(parse_date) if "执行时间" in df else pd.NaT
    df["__is_report"] = df.get("__is_report", False) if "__is_report" in df.columns else False
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
    df["__store_raw"] = df[store_col].map(clean_text) if store_col else ""
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
    month_columns = [column for column in sales_df.columns if re.fullmatch(r"20\d{2}-\d{2}", str(column))]
    sales_df["__sale_month"] = sales_df["__sale_date"].dt.to_period("M")
    sales_df["__in_month"] = (
        sales_df["__sale_month"].eq(target_period)
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
            "monthly_qty": {},
            "sale_dates": [],
        }
        if pd.notna(row.get("__sale_month")):
            record["monthly_qty"][row["__sale_month"]] = float(row.get("__qty", 0) or 0)
        for month_column in month_columns:
            qty = pd.to_numeric(pd.Series([row.get(month_column)]), errors="coerce").fillna(0).iloc[0]
            if qty:
                record["monthly_qty"][pd.Period(str(month_column), freq="M")] = float(qty)
        if pd.notna(row.get("__sale_date")):
            record["sale_dates"].append(row.get("__sale_date"))
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
                    for month, qty in record["monthly_qty"].items():
                        old["monthly_qty"][month] = old["monthly_qty"].get(month, 0) + qty
                    old["sale_dates"].extend(record["sale_dates"])
                    if pd.notna(record["last_sale_date"]):
                        old["last_sale_date"] = record["last_sale_date"]
                        old["last_qty"] = record["last_qty"]
    for record in lookup.values():
        record["should_have_no_purchase_reason"] = should_have_no_purchase_reason(record, target_month)
    return lookup


def find_sale_by_name_store(row: pd.Series, lookup: dict[str, dict]) -> dict:
    key = clean_text(row.get("__key_name_store"))
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


def should_have_no_purchase_reason_from_follow(row: pd.Series, target_month: str) -> str:
    """Determine stock status from follow-up data (not sales lookup).
    Returns '否' if patient has stock (should NOT have purchase reason),
    '是' if patient has no stock (should HAVE purchase reason).
    """
    last_purchase_date = parse_date(row.get("会员最近一次门店购药时间"))
    last_purchase_boxes = pd.to_numeric(row.get("会员最近一次门店购药盒数"), errors="coerce")

    target = pd.Period(target_month, freq="M")

    if pd.isna(last_purchase_date):
        # No purchase info → assume no stock
        return "是"

    last_month = last_purchase_date.to_period("M")

    # If last purchase is in the target month → purchased this month → has stock
    if last_month == target:
        return "否"

    # Calculate remaining stock from last purchase
    # Also consider second purchase for cross-month accumulation
    months_elapsed = (target - last_month).n
    remaining = (last_purchase_boxes if pd.notna(last_purchase_boxes) else 0) - months_elapsed

    # Consider 倒数第二次门店购药 for cross-month accumulation
    second_purchase_date = parse_date(row.get("倒数第二次门店购药时间"))
    second_purchase_boxes = pd.to_numeric(row.get("会员倒数第二次门店购药盒数"), errors="coerce")
    if pd.notna(second_purchase_date) and pd.notna(second_purchase_boxes):
        second_month = second_purchase_date.to_period("M")
        # Add second purchase boxes, but only if they contribute to current stock
        second_elapsed = (target - second_month).n
        second_remaining = second_purchase_boxes - second_elapsed
        if second_remaining > 0:
            remaining += second_remaining

    if remaining > 0:
        return "否"
    return "是"


def should_have_from_sales(monthly_qty: dict, target_month_str: str) -> str:
    """Rolling-window stock determination for 泰瑞沙/利普卓 (automation formula).
    Returns '否' if patient has stock (should NOT have purchase reason),
    '是' if patient has no stock (should HAVE purchase reason).

    Logic: if target month has purchase → '否';
    for k=1..21: if sum(months target-k to target) > k → '否' (has stock);
    else → '是' (should have reason).
    """
    target = pd.Period(target_month_str, freq="M")
    if monthly_qty.get(target, 0) > 0:
        return "否"
    for k in range(1, 22):
        window_sum = sum(monthly_qty.get(target - offset, 0) for offset in range(k + 1))
        if window_sum > k:
            return "否"
    return "是"


def compute_tr_reason(follow_row: pd.Series, should_value: str) -> str:
    """泰瑞沙 匹配结果 formula: TRIM(concat of indication + purchase parts).
    If empty → no issue → row is NOT included in output.
    """
    indication = clean_text(follow_row.get("适应症"))
    no_purchase = clean_text(follow_row.get("本月未购药的原因"))

    indication_part = ""
    purchase_part = ""

    if not indication:
        indication_part = "适应症未填写需要填写适应症"

    if should_value == "是" and not no_purchase:
        purchase_part = "患者本月还未来购药且没有存药，需填写本月未购药原因"
    elif should_value == "否" and no_purchase:
        purchase_part = "患者本月已来购药，本月未购药原因应该为空"

    if indication_part and purchase_part:
        return indication_part + ", " + purchase_part
    return indication_part or purchase_part


def first_follow_match(follow_df: pd.DataFrame, sale_row: pd.Series) -> pd.Series:
    phone = clean_text(sale_row.get("__phone"))
    member = clean_text(sale_row.get("__member"))
    name_store = clean_text(sale_row.get("__key_name_store"))
    if phone:
        matched = follow_df[follow_df["__phone"] == phone]
        if not matched.empty:
            return matched.iloc[0]
    if member:
        matched = follow_df[follow_df["__member"] == member]
        if not matched.empty:
            return matched.iloc[0]
    if name_store:
        matched = follow_df[follow_df["__key_name_store"] == name_store]
        if not matched.empty:
            return matched.iloc[0]
    return pd.Series(dtype=object)


def build_output_row(brand: str, source_row: pd.Series, follow_row: pd.Series, reason: str,
                     delay_label: str = "", no_purchase_label: str = "") -> dict:
    row = follow_row if not follow_row.empty else source_row
    # For 优赫得/凡舒卓, use computed labels; for others, use raw follow-up values
    if delay_label or no_purchase_label:
        delay_value = delay_label
        no_purchase_value = no_purchase_label
    else:
        delay_value = clean_text(row.get("患者延迟用药的原因"))
        no_purchase_value = clean_text(row.get("本月未购药的原因"))

    # Determine store name format based on brand:
    # 泰瑞沙/利普卓/英飞凡: correct list uses RAW (full) store names → prefer follow data "门店"
    # 荃科得/优赫得/凡舒卓: correct list uses ALIASED (short) store names → prefer sale data "药房名称"
    store_value = ""
    if brand in RAW_STORE_BRANDS:
        # RAW brands: prefer follow data's __store_raw (full name from "门店")
        store_value = clean_text(row.get("__store_raw")) or clean_text(row.get("门店"))
        if store_value in REVERSE_ALIASES:
            store_value = REVERSE_ALIASES[store_value]
        if not store_value:
            store_value = clean_text(source_row.get("__store_raw")) or clean_text(source_row.get("__store"))
            if store_value in REVERSE_ALIASES:
                store_value = REVERSE_ALIASES[store_value]
    else:
        # ALIASED brands: prefer SALE data's __store (aliased form from "药房名称")
        # This ensures the store name matches where the patient actually purchased
        store_value = clean_text(source_row.get("__store"))
        if not store_value:
            store_value = clean_text(row.get("__store"))
        if not store_value:
            store_value = clean_text(row.get("__store_raw")) or clean_text(row.get("门店"))
            if store_value in PHARMACY_ALIASES:
                store_value = PHARMACY_ALIASES[store_value]

    return {
        "品牌": brand,
        "药房": store_value,
        "执行时间": row.get("执行时间", ""),
        "患者姓名": clean_text(row.get("患者姓名")) or clean_text(source_row.get("会员姓名")) or clean_text(source_row.get("开票抬头")),
        "联系方式": clean_text(row.get("患者手机号")) or clean_text(source_row.get("会员电话")),
        "适应症": clean_text(row.get("适应症")) or clean_text(source_row.get("适应症")),
        "会员最近一次门店购药时间": row.get("会员最近一次门店购药时间", source_row.get("销售时间", "")),
        "会员最近一次门店购药盒数": row.get("会员最近一次门店购药盒数", source_row.get("销售数量", "")),
        "延期用药原因": delay_value,
        "本月未购药的原因": no_purchase_value,
        "医生建议用药时长": clean_text(row.get("医生建议用药时长")),
        "需要补充的原因": reason,
    }


def target_period(target_month: str) -> pd.Period:
    return pd.Period(target_month, freq="M")


def in_target_month(value, target_month: str) -> bool:
    parsed = parse_date(value)
    return pd.notna(parsed) and parsed.to_period("M") == target_period(target_month)


def sale_patient_key(row: pd.Series) -> str:
    key = clean_text(row.get("__key_name_store"))
    if key and not key.endswith("|"):
        return key
    phone_key = clean_text(row.get("__key_phone"))
    if phone_key and not phone_key.endswith("|"):
        return phone_key
    member_key = clean_text(row.get("__key_member"))
    if member_key and not member_key.endswith("|"):
        return member_key
    return ""


def person_key(row: pd.Series) -> str:
    brand = clean_text(row.get("__brand"))
    phone = clean_text(row.get("__phone"))
    name = clean_text(row.get("__name")) or clean_text(row.get("患者姓名")) or clean_text(row.get("会员姓名")) or clean_text(row.get("开票抬头"))
    if phone:
        return brand + "|phone|" + phone
    if name:
        return brand + "|name|" + name
    return ""


def find_follow_by_sale(follow_df: pd.DataFrame, sale_row: pd.Series) -> pd.Series:
    key = clean_text(sale_row.get("__key_name_store"))
    if key:
        matched = follow_df[follow_df["__key_name_store"] == key]
        if not matched.empty:
            return matched.sort_values("__date", ascending=False).iloc[0]
    return first_follow_match(follow_df, sale_row)


def find_follow_by_person(follow_df: pd.DataFrame, sale_row: pd.Series, target_month: str) -> pd.Series:
    candidate = follow_df[follow_df["__date"].map(lambda value: in_target_month(value, target_month))].copy()
    phone = clean_text(sale_row.get("__phone"))
    name = clean_text(sale_row.get("__name"))
    if phone:
        matched = candidate[candidate["__phone"].eq(phone)]
        if not matched.empty:
            return matched.sort_values("__date", ascending=False).iloc[0]
    if name:
        matched = candidate[candidate["__name"].eq(name)]
        if not matched.empty:
            return matched.sort_values("__date", ascending=False).iloc[0]
    return pd.Series(dtype=object)


def has_target_purchase(history: pd.DataFrame, target_month: str) -> bool:
    if history.empty:
        return False
    dates = history["__sale_date"].map(parse_date)
    return dates.dt.to_period("M").eq(target_period(target_month)).any()


def rule_stock_brand_follow(row: pd.Series, target_month: str) -> str:
    """泰瑞沙/利普卓 rule: check 本月未购药原因 based on follow-up stock status."""
    indication = clean_text(row.get("适应症"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    should_have_reason = should_have_no_purchase_reason_from_follow(row, target_month)
    parts = []
    if not indication:
        parts.append("适应症未填写需要填写适应症")
    if should_have_reason == "是" and not no_purchase_reason:
        parts.append("患者本月还未来购药且没有存药，需填写本月未购药原因")
    if should_have_reason == "否" and no_purchase_reason:
        parts.append("患者本月已来购药，本月未购药原因应该为空")
    return ", ".join(parts)


def rule_yingfeifan_patient(sale_row: pd.Series, follow_row: pd.Series, history: pd.DataFrame, target_month: str) -> str:
    bought_this_month = has_target_purchase(history, target_month)
    if follow_row.empty:
        return "本月患者还未随访"

    no_purchase_reason = clean_text(follow_row.get("本月未购药的原因"))
    follow_purchase = yes_value(follow_row.get("当月是否购药"))
    indication = clean_text(follow_row.get("适应症"))

    parts = []
    if bought_this_month and not (follow_purchase is True and not no_purchase_reason):
        parts.append("本月患者已购药，请检查当月是否购药列和本月未购药原因是否准确")
    if not bought_this_month and not (follow_purchase is False and no_purchase_reason):
        parts.append("本月患者未购药，请检查当月是否购药列和本月未购药原因是否准确")
    if not indication:
        parts.append("适应症为空需重新生成随访进行填写")
    return "；".join(parts)


def compute_yh_issue(expected_dt, follow_q, follow_v, follow_o):
    """优赫得 AE formula (per-transaction): returns (issue_reason, AD_label, AC_label).
    expected_dt preserves time component so that expected=2026-06-29 09:01 > TODAY=2026-06-29 00:00."""
    TODAY = pd.Timestamp.today().normalize()
    C = expected_dt

    # AA: actual purchase date logic
    if C > TODAY:
        # Expected in future: if follow-up purchase is near expected (±5 days), use it; else "日期未到"
        if pd.notna(follow_q) and follow_q.normalize() >= (C - pd.Timedelta(days=5)).normalize():
            AA = follow_q.normalize()
        else:
            AA = "日期未到"
    else:
        # Expected has passed: use follow-up purchase date if available
        if pd.notna(follow_q):
            AA = follow_q.normalize()
        else:
            AA = pd.NaT

    # AB: 是否超期
    if isinstance(AA, str) and AA == "日期未到":
        AB = "规律"
    elif pd.isna(AA):
        AB = "超期"
    else:
        AA_date = pd.Timestamp(AA)
        C_norm = C.normalize()
        if AA_date > C_norm or AA_date < (C_norm - pd.Timedelta(days=5)):
            AB = "超期"
        else:
            AB = "规律"

    # AC: 本月未购药原因 assessment
    v = follow_v
    if AB == "规律":
        AC = "" if not v else "错误：规律但存在未购药原因"
    elif AB == "超期":
        AC = v if v else "错误：超期但无未购药原因"
    else:
        AC = ""

    # AD: 延期用药原因 assessment
    o = follow_o
    if AB == "规律":
        AD = "" if not o else "错误：规律但存在延期用药原因"
    elif AB == "超期":
        AD = o if o else "错误：超期但无延期用药原因"
    else:
        AD = ""

    # AE: final issue determination
    if C > TODAY and isinstance(AA, str) and AA == "日期未到":
        return (f"预计{C.normalize().strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因", AD, AC)

    ac_is_error = (not AC) or ("错误" in AC)
    ad_is_error = (not AD) or ("错误" in AD)

    if AB == "超期" and ac_is_error and ad_is_error:
        return ("已超期但未记录未购药原因或延期用药原因，需补充", AD, AC)

    if not pd.isna(AA) and not isinstance(AA, str):
        AA_date = pd.Timestamp(AA)
        C_norm = C.normalize()
        if (AA_date > C_norm and AA_date > pd.Timestamp(0) and C_norm > pd.Timestamp(0) and
            AA_date.month == C_norm.month and AA_date.year == C_norm.year):
            if AD and "错误" not in AD:
                return ("", AD, AC)
            return ("患者延迟用药，应填写延迟用药原因", AD, AC)

    if AC == "错误：规律但存在未购药原因":
        return ("需填写本月未购药原因/延期用药原因", AD, AC)

    if AC == "错误：超期但无未购药原因":
        return ("待观察，如果本月内未来购药，需重新生成随访填写本月未购药原因", AD, AC)

    if AB == "规律" and AC:
        return ("购药规律但记录了未购药原因，需核实", AD, AC)

    return ("", AD, AC)


def compute_fsz_issue(F, H, I_valid, follow_v):
    """凡舒卓 M formula (per-patient) WITH Step 0: returns (issue_reason, L_label).
    F = last sale date (normalized), H = expected date (with time), I_valid = whether phone matched in follow-up."""
    TODAY = pd.Timestamp.today().normalize()

    # J: actual purchase date logic
    if pd.notna(H) and H > TODAY:
        J = "日期未到"
    else:
        J = F

    # K: 是否超期
    if isinstance(J, str) and J == "日期未到":
        K = "规律"
    elif pd.isna(J):
        K = "超期"
    else:
        J_date = pd.Timestamp(J) if not isinstance(J, pd.Timestamp) else J
        H_norm = H.normalize() if pd.notna(H) else pd.NaT
        if pd.isna(H_norm):
            K = "超期"
        elif J_date > H_norm or J_date < (H_norm - pd.Timedelta(days=5)):
            K = "超期"
        else:
            K = "规律"

    # L: 未购药原因 assessment
    v = follow_v
    if K == "规律":
        L = "" if not v else "错误：规律但存在填写了未购药原因"
    elif K == "超期":
        L = v if v else "错误：超期但无未购药原因"
    else:
        L = ""

    # ─── Step 0 (CRITICAL!) ───
    # IF patient bought this month AND expected next month AND has followup AND no issue → ""
    eom_today = (TODAY + pd.offsets.MonthEnd(0)).normalize()
    bom_today = (TODAY - pd.offsets.MonthBegin(1) + pd.Timedelta(days=1)).normalize()

    if (pd.notna(H) and H > eom_today and
        pd.notna(F) and F >= bom_today and F <= eom_today and
        I_valid and not L):
        return ("", L)

    # Step 1: No follow-up match
    if not I_valid:
        if K == "超期":
            return (f"应做随访未做随访，该患者预计购药日期为：{H.normalize().strftime('%Y-%m-%d')}，目前已经超期，已超期但未记录本月未购药原因，需补充", L)
        return (f"应做随访未做随访，该患者预计购药日期为：{H.normalize().strftime('%Y-%m-%d')}", L)

    # Step 2: Expected > TODAY AND J = "日期未到"
    if pd.notna(H) and H > TODAY and isinstance(J, str) and J == "日期未到":
        return (f"预计{H.normalize().strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因", L)

    # Step 3: K = "超期" AND (L empty or "错误")
    l_is_error = (not L) or ("错误" in L)
    if K == "超期" and l_is_error:
        if not isinstance(J, str) and pd.notna(J):
            J_date = pd.Timestamp(J) if not isinstance(J, pd.Timestamp) else J
            H_norm = H.normalize() if pd.notna(H) else pd.NaT
            if (pd.notna(H_norm) and J_date > H_norm and
                J_date > pd.Timestamp(0) and H_norm > pd.Timestamp(0) and
                J_date.month == H_norm.month and J_date.year == H_norm.year):
                return ("患者目前已经超期，但随访未记录本月未购药原因，需补充", L)
        return ("已超期但未记录本月未购药原因，需补充", L)

    # Step 5: L = "错误：规律但存在未购药原因"
    if L == "错误：规律但存在未购药原因":
        return ("需填写本月未购药原因", L)

    # Step 6: L = "错误：超期但无未购药原因"
    if L == "错误：超期但无未购药原因":
        return ("待观察，如果本月内未来购药，需重新生成随访填写本月未购药原因", L)

    # Step 7: K = "规律" AND L <> ""
    if K == "规律" and L:
        return ("购药规律但记录了未购药原因，需核实", L)

    return ("", L)


def quankede_reason_for_patient(follow_rows: pd.DataFrame, last_sale_row: pd.Series, target_month: str) -> str:
    if follow_rows.empty or "__date" not in follow_rows.columns:
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    target_rows = follow_rows[follow_rows["__date"].map(lambda value: in_target_month(value, target_month))].copy()
    dates = target_rows["__date"].map(parse_date).dropna().sort_values(ascending=False)
    if len(dates) == 0:
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    if len(dates) == 1:
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    days = abs((dates.iloc[0] - dates.iloc[1]).days)
    if days <= 10:
        return f"本月已随访两次，但两次间隔＜十天，目前间隔 {days} 天，请重新生成随访"
    return ""


def build_issue_table(follow_df: pd.DataFrame, sales_df: pd.DataFrame, target_month: str) -> pd.DataFrame:
    follow_df = add_follow_keys(follow_df)
    sales_df = add_sales_keys(sales_df, target_month) if not sales_df.empty else pd.DataFrame()
    follow_brands = [brand for brand in follow_df["__brand"].dropna().unique() if clean_text(brand)]
    if sales_df.empty:
        sales_df = pd.DataFrame(columns=["__brand", "__phone", "__member", "__name", "__store", "__key_phone", "__key_member", "__key_name_store"])
    if len(follow_brands) == 1 and not sales_df.empty:
        sales_df.loc[sales_df["__brand"].map(clean_text).eq(""), "__brand"] = follow_brands[0]
        sales_df["__key_phone"] = sales_df["__brand"] + "|" + sales_df["__phone"]
        sales_df["__key_member"] = sales_df["__brand"] + "|" + sales_df["__member"]
        sales_df["__key_name_store"] = sales_df["__brand"] + "|" + sales_df["__name"] + "|" + sales_df["__store"]

    rows = []

    # ─── 泰瑞沙 ───
    # Rolling-window approach: use SALES data monthly totals to determine stock,
    # then XLOOKUP into follow-up data by name+short_store key.
    # Scope filter: only include follow patients at 15 mapped pharmacies.
    tr_sales = sales_df[sales_df["__brand"].eq("泰瑞沙")].copy() if not sales_df.empty else pd.DataFrame()
    if not tr_sales.empty:
        # Build monthly sales pivot: aggregate qty per (name|short_store) per month
        tr_sales["__sale_month"] = tr_sales["__sale_date"].dt.to_period("M")
        tr_sales["__qty"] = pd.to_numeric(tr_sales.get("销售数量", 1), errors="coerce").fillna(1)
        tr_sales["__sale_key"] = tr_sales["__name"] + "|" + tr_sales["__store_raw"]  # raw short store from sales (avoid normalize_store bracket mismatch)
        monthly_pivot = tr_sales.groupby(["__sale_key", "__sale_month"])["__qty"].sum().reset_index()
        monthly_dict: dict[str, dict] = {}
        for _, piv_row in monthly_pivot.iterrows():
            key = clean_text(piv_row.get("__sale_key"))
            month = piv_row.get("__sale_month")
            qty = float(piv_row.get("__qty", 0))
            if key not in monthly_dict:
                monthly_dict[key] = {}
            monthly_dict[key][month] = monthly_dict[key].get(month, 0) + qty

        # Apply rolling window formula to each sales key
        should_lookup = {key: should_have_from_sales(monthly, target_month) for key, monthly in monthly_dict.items()}

        # Prepare follow-up data with SCOPE filter
        tr_follow = follow_df[follow_df["__brand"].eq("泰瑞沙")].copy()
        if not tr_follow.empty:
            # Map 门店 (full name) → short name via PHARMACY_SCOPE; only keep in-scope rows
            tr_follow["__store_short"] = tr_follow["门店"].map(
                lambda v: PHARMACY_SCOPE.get(clean_text(v), None)
            )
            tr_follow_scope = tr_follow[tr_follow["__store_short"].notna()].copy()
            # Matching key: name + short_store (from scope mapping)
            tr_follow_scope["__follow_key"] = tr_follow_scope["__name"] + "|" + tr_follow_scope["__store_short"]
            # Dedup: keep most recent per name+store, prefer report rows
            tr_follow_sorted = tr_follow_scope.sort_values(
                ["__is_report", "__date"], ascending=[False, False]
            )
            tr_follow_dedup = tr_follow_sorted.drop_duplicates(subset=["__name", "__store_short"])

            for _, follow_row in tr_follow_dedup.iterrows():
                if not clean_text(follow_row.get("患者姓名")):
                    continue
                # XLOOKUP: match by name+short_store key
                follow_key = clean_text(follow_row.get("__follow_key"))
                should_value = should_lookup.get(follow_key)
                if should_value is None:
                    # No sales match → skip (patient has no purchase history in scope stores)
                    continue
                reason = compute_tr_reason(follow_row, should_value)
                if not reason:
                    continue
                # Store name: RAW format (full name from 门店 column)
                rows.append(build_output_row("泰瑞沙", pd.Series(dtype=object), follow_row, reason))

    # ─── 利普卓 ───
    # Same rolling-window approach as 泰瑞沙: use SALES data monthly totals,
    # then XLOOKUP into follow-up data by name+short_store key.
    # Scope filter: only include follow patients at 15 mapped pharmacies.
    lp_sales = sales_df[sales_df["__brand"].eq("利普卓")].copy() if not sales_df.empty else pd.DataFrame()
    if not lp_sales.empty:
        lp_sales["__sale_month"] = lp_sales["__sale_date"].dt.to_period("M")
        lp_sales["__qty"] = pd.to_numeric(lp_sales.get("销售数量", 1), errors="coerce").fillna(1)
        lp_sales["__sale_key"] = lp_sales["__name"] + "|" + lp_sales["__store_raw"]
        lp_monthly_pivot = lp_sales.groupby(["__sale_key", "__sale_month"])["__qty"].sum().reset_index()
        lp_monthly_dict: dict[str, dict] = {}
        for _, piv_row in lp_monthly_pivot.iterrows():
            key = clean_text(piv_row.get("__sale_key"))
            month = piv_row.get("__sale_month")
            qty = float(piv_row.get("__qty", 0))
            if key not in lp_monthly_dict:
                lp_monthly_dict[key] = {}
            lp_monthly_dict[key][month] = lp_monthly_dict[key].get(month, 0) + qty

        lp_should_lookup = {key: should_have_from_sales(monthly, target_month) for key, monthly in lp_monthly_dict.items()}

        lp_follow = follow_df[follow_df["__brand"].eq("利普卓")].copy()
        if not lp_follow.empty:
            lp_follow["__store_short"] = lp_follow["门店"].map(
                lambda v: PHARMACY_SCOPE.get(clean_text(v), None)
            )
            lp_follow_scope = lp_follow[lp_follow["__store_short"].notna()].copy()
            lp_follow_scope["__follow_key"] = lp_follow_scope["__name"] + "|" + lp_follow_scope["__store_short"]
            lp_follow_sorted = lp_follow_scope.sort_values(
                ["__is_report", "__date"], ascending=[False, False]
            )
            lp_follow_dedup = lp_follow_sorted.drop_duplicates(subset=["__name", "__store_short"])

            for _, follow_row in lp_follow_dedup.iterrows():
                if not clean_text(follow_row.get("患者姓名")):
                    continue
                follow_key = clean_text(follow_row.get("__follow_key"))
                should_value = lp_should_lookup.get(follow_key)
                if should_value is None:
                    continue
                reason = compute_tr_reason(follow_row, should_value)
                if not reason:
                    continue
                rows.append(build_output_row("利普卓", pd.Series(dtype=object), follow_row, reason))

    # ─── 优赫得 ───
    # Per-transaction approach (AE formula): each sale row with expected date in target month
    # is independently processed. Follow-up matched by PHONE, not name+store.
    # Pharmacy scope: only 南充药房 and 德阳关爱药房.
    yh_sales = sales_df[
        sales_df["__brand"].eq("优赫得")
        & sales_df["__store"].isin(["南充药房", "德阳关爱药房"])
    ].copy()
    yh_follow = follow_df[follow_df["__brand"].eq("优赫得")].copy()

    if not yh_sales.empty:
        # Add datetime-preserved sale date for expected computation
        yh_sales["__sale_date_dt"] = yh_sales["销售时间"].map(parse_datetime)
        yh_sales["__expected"] = yh_sales["__sale_date_dt"] + pd.Timedelta(days=21)

        # Filter: only rows where expected date falls in target month
        tp = target_period(target_month)
        yh_filtered = yh_sales[
            yh_sales["__expected"].map(lambda x: pd.notna(x) and x.to_period("M") == tp)
        ].copy()

        # Build phone-based follow-up lookup (keep most recent per phone, prefer report)
        yh_follow_by_phone: dict[str, pd.Series] = {}
        for _, frow in yh_follow.iterrows():
            phone = clean_text(frow.get("__phone"))
            if phone:
                if phone not in yh_follow_by_phone:
                    yh_follow_by_phone[phone] = frow
                else:
                    existing = yh_follow_by_phone[phone]
                    # Prefer report rows; then prefer later dates
                    existing_is_report = existing.get("__is_report", False)
                    frow_is_report = frow.get("__is_report", False)
                    existing_date = parse_date(existing.get("执行时间"))
                    frow_date = parse_date(frow.get("执行时间"))
                    if (frow_is_report and not existing_is_report) or \
                       (frow_is_report == existing_is_report and pd.notna(frow_date) and (pd.isna(existing_date) or frow_date > existing_date)):
                        yh_follow_by_phone[phone] = frow

        # Process each transaction row independently
        yh_results = []
        for _, sale_row in yh_filtered.iterrows():
            phone = clean_text(sale_row.get("__phone"))
            expected = sale_row["__expected"]
            follow_row = yh_follow_by_phone.get(phone, pd.Series(dtype=object))
            follow_q = parse_date(follow_row.get("会员最近一次门店购药时间")) if not follow_row.empty else pd.NaT
            follow_v = clean_text(follow_row.get("本月未购药的原因")) if not follow_row.empty else ""
            follow_o = clean_text(follow_row.get("患者延迟用药的原因")) if not follow_row.empty else ""

            issue, ad, ac = compute_yh_issue(expected, follow_q, follow_v, follow_o)
            if issue:
                yh_results.append({
                    "sale_row": sale_row,
                    "follow_row": follow_row if not follow_row.empty else pd.Series(dtype=object),
                    "issue": issue,
                    "ad": ad,
                    "ac": ac,
                    "phone": phone,
                    "name": clean_text(sale_row.get("__name")),
                    "store": clean_text(sale_row.get("__store")),
                })

        # Dedup by (phone, name, store, reason)
        yh_seen = set()
        for r in yh_results:
            dedup_key = (r["phone"], r["name"], r["store"], r["issue"])
            if dedup_key in yh_seen:
                continue
            yh_seen.add(dedup_key)
            rows.append(build_output_row("优赫得", r["sale_row"], r["follow_row"], r["issue"], r["ad"], r["ac"]))

    # ─── 凡舒卓 ───
    # Per-patient approach (M formula WITH Step 0): group by name+store,
    # use LATEST sale per group. Follow-up matched by PHONE for I_valid check.
    # No pharmacy scope filter (all stores).
    fsz_sales = sales_df[sales_df["__brand"].eq("凡舒卓")].copy()
    fsz_follow = follow_df[follow_df["__brand"].eq("凡舒卓")].copy()

    if not fsz_sales.empty:
        # Build phone-based follow-up lookup for I_valid check
        fsz_follow_by_phone: dict[str, pd.Series] = {}
        for _, frow in fsz_follow.iterrows():
            phone = clean_text(frow.get("__phone"))
            if phone:
                if phone not in fsz_follow_by_phone:
                    fsz_follow_by_phone[phone] = frow
                else:
                    existing = fsz_follow_by_phone[phone]
                    existing_is_report = existing.get("__is_report", False)
                    frow_is_report = frow.get("__is_report", False)
                    existing_date = parse_date(existing.get("执行时间"))
                    frow_date = parse_date(frow.get("执行时间"))
                    if (frow_is_report and not existing_is_report) or \
                       (frow_is_report == existing_is_report and pd.notna(frow_date) and (pd.isna(existing_date) or frow_date > existing_date)):
                        fsz_follow_by_phone[phone] = frow

        # Per-patient: group by name+store, use latest sale
        fsz_sales["__sale_key"] = fsz_sales["__name"] + "|" + fsz_sales["__store"]
        fsz_processed = set()
        for sale_key, group in fsz_sales.groupby("__sale_key"):
            if not sale_key or sale_key in fsz_processed:
                continue
            fsz_processed.add(sale_key)

            # Use latest sale record
            latest = group.sort_values("__sale_date", ascending=False).iloc[0]
            sale_date_dt = parse_datetime(latest.get("销售时间"))
            sale_date_norm = parse_date(latest.get("销售时间"))
            if pd.isna(sale_date_dt):
                continue

            # F = normalized sale date, H = sale_date + 28 days (preserving time)
            F = sale_date_norm
            H = sale_date_dt + pd.Timedelta(days=28)

            # Broad expected date filter: Jan to Aug 2026
            expected_period = H.to_period("M")
            tp = target_period(target_month)
            if expected_period < tp - 5 or expected_period > tp + 1:
                continue

            # Phone-based I_valid check (XLOOKUP phone in follow data)
            phone = clean_text(latest.get("__phone"))
            I_valid = bool(phone) and phone in fsz_follow_by_phone
            follow_row = fsz_follow_by_phone.get(phone, pd.Series(dtype=object))
            follow_v = clean_text(follow_row.get("本月未购药的原因")) if not follow_row.empty else ""

            issue, L = compute_fsz_issue(F, H, I_valid, follow_v)
            if issue:
                rows.append(build_output_row("凡舒卓", latest, follow_row, issue, "", L))

    # ─── 荃科得 ───
    # ALL purchase patients this month must have 2 follow-ups with >10 day gap
    qkd_sales = sales_df[sales_df["__brand"].eq("荃科得")].copy()
    qkd_follow = follow_df[follow_df["__brand"].eq("荃科得")].copy()
    processed_keys = set()
    # Get all unique patients from sales (any time, to also catch those without follow-up)
    for key, history in qkd_sales.groupby(qkd_sales.apply(sale_patient_key, axis=1)):
        if not key:
            continue
        processed_keys.add(key)
        follow_rows = qkd_follow[qkd_follow["__key_name_store"].eq(key)].copy()
        sale_row = history.sort_values("__sale_date").iloc[-1]
        if follow_rows.empty:
            follow_row = find_follow_by_sale(qkd_follow, sale_row)
        else:
            follow_row = follow_rows.sort_values("__date", ascending=False).iloc[0]
        reason = quankede_reason_for_patient(follow_rows, sale_row, target_month)
        if reason:
            rows.append(build_output_row("荃科得", sale_row, follow_row, reason))

    # Also check 荃科得 patients in follow-up who may not be in sales
    # Use broader matching: phone, member, name+store keys
    qkd_follow_dedup = qkd_follow.sort_values("__date", ascending=False).drop_duplicates(subset=["__name", "__store"])
    for _, follow_row in qkd_follow_dedup.iterrows():
        # Skip if already processed from sales (check by name+store)
        name = clean_text(follow_row.get("患者姓名"))
        store = clean_text(follow_row.get("__store"))
        brand_key = "荃科得|" + name + "|" + store
        if brand_key in processed_keys:
            continue
        # Also check by phone and member keys
        phone_key = "荃科得|" + clean_text(follow_row.get("__phone"))
        member_key = "荃科得|" + clean_text(follow_row.get("__member"))
        if phone_key in processed_keys or member_key in processed_keys:
            continue
        key = clean_text(follow_row.get("__key_name_store"))
        if not key:
            continue
        follow_rows = qkd_follow[qkd_follow["__key_name_store"].eq(key)].copy()
        # Also try matching by phone
        if follow_rows.empty and phone_key and not phone_key.endswith("|"):
            follow_rows = qkd_follow[qkd_follow["__key_phone"].eq(phone_key)].copy()
        reason = quankede_reason_for_patient(follow_rows, pd.Series(dtype=object), target_month)
        if reason:
            rows.append(build_output_row("荃科得", pd.Series(dtype=object), follow_row, reason))

    # ─── 英飞凡 ───
    # 晟德药房 patients with 2025+ purchase records
    # Per-patient dedup, use person_key for grouping, then match follow by person
    yingfei_sales = sales_df[
        sales_df["__brand"].eq("英飞凡")
        & sales_df["__store"].map(clean_text).str.contains("晟德", na=False)
        & sales_df["__sale_date"].map(lambda value: pd.notna(parse_date(value)) and parse_date(value) >= pd.Timestamp("2025-01-01"))
    ].copy()
    yingfei_follow = follow_df[follow_df["__brand"].eq("英飞凡")].copy()
    yingfei_processed = set()
    for key, history in yingfei_sales.groupby(yingfei_sales.apply(person_key, axis=1)):
        if not key:
            continue
        if key in yingfei_processed:
            continue
        yingfei_processed.add(key)
        sale_row = history.sort_values("__sale_date").iloc[-1]
        follow_row = find_follow_by_person(yingfei_follow, sale_row, target_month)
        reason = rule_yingfeifan_patient(sale_row, follow_row, history, target_month)
        if reason:
            rows.append(build_output_row("英飞凡", sale_row, follow_row, reason))

    result = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if not result.empty:
        result = result.drop_duplicates()
        result = result.sort_values(["品牌", "药房", "患者姓名"], na_position="last")
    return result


def to_excel_bytes(result: pd.DataFrame, follow_df: pd.DataFrame, sales_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="问题表格清单")
    return output.getvalue()


st.title("DTP 自动化问题清单")
st.caption("上传销售底表和随访底表，自动合并生成所有品种的问题表格清单。")

with st.sidebar:
    st.header("使用方式")
    st.markdown("1. 上传销售底表")
    st.markdown("2. 上传一个或多个随访底表")
    st.markdown("3. 选择统计月份")
    st.markdown("4. 下载问题表格清单")

sales_file = st.file_uploader("销售底表", type=["xlsx", "xls"], key="sales")
follow_files = st.file_uploader(
    "随访底表（可一次上传多个）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="follow",
    help="可以同时上传优赫得/荃科得专属随访表和其他品种通用随访表，系统会按品牌自动分流。",
)

if not sales_file or not follow_files:
    st.info("请先上传销售底表，并至少上传一个随访底表。")
    st.write("当前支持泰瑞沙、利普卓、英飞凡、荃科得、优赫得、凡舒卓等品种字段。")
    st.write("如果上传多份随访底表，优赫得会自动走专属表，其他品种走通用表。")
    st.stop()

try:
    sales_sheets = read_excel_upload(sales_file)
    sales_df = pick_sheet(sales_sheets, "销售底表")
    follow_df = combine_follow_files(follow_files)
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
