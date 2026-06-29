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


def has_on_time_purchase(history: pd.DataFrame, expected_date) -> bool:
    if history.empty or pd.isna(expected_date):
        return False
    dates = history["__sale_date"].map(parse_date)
    return dates.between(expected_date - pd.Timedelta(days=5), expected_date, inclusive="both").any()


def has_later_target_purchase(history: pd.DataFrame, expected_date, target_month: str) -> bool:
    if history.empty or pd.isna(expected_date):
        return False
    dates = history["__sale_date"].map(parse_date)
    return ((dates > expected_date) & dates.dt.to_period("M").eq(target_period(target_month))).any()


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


def transaction_actual_purchase(expected_date, follow_row: pd.Series, history: pd.DataFrame = None):
    """Determine actual purchase date relative to expected date.
    If history is provided (for 优赫得/凡舒卓), use sales data to find purchase near expected date.
    Otherwise, use follow-up data's 会员最近一次门店购药时间.
    """
    if pd.isna(expected_date):
        return pd.NaT

    # For 优赫得/凡舒卓: use sales history to find purchase near expected date
    if history is not None and not history.empty:
        dates = history["__sale_date"].map(parse_date).dropna()
        # Find purchase within 5 days before expected date or on expected date
        near_expected = dates[
            dates.between(expected_date - pd.Timedelta(days=5), expected_date + pd.Timedelta(days=5), inclusive="both")
        ]
        if not near_expected.empty:
            return near_expected.iloc[-1]  # Most recent near-expected purchase
        # No purchase near expected date → check if expected is in the future
        today = pd.Timestamp.today().normalize()
        if expected_date > today:
            return "日期未到"
        # Expected date has passed but no purchase near it → overdue
        return pd.NaT  # No purchase = overdue

    # Original logic for 泰瑞沙/利普卓/英飞凡: use follow-up data
    follow_recent = parse_date(follow_row.get("会员最近一次门店购药时间"))
    today = pd.Timestamp.today().normalize()
    if expected_date <= today:
        return follow_recent
    if pd.notna(follow_recent) and follow_recent >= expected_date - pd.Timedelta(days=5):
        return follow_recent
    return "日期未到"


def purchase_status_label(actual_purchase, expected_date) -> str:
    """Classify purchase timing: 规律 (on time), 超期 (overdue), or empty."""
    if clean_text(actual_purchase) == "日期未到":
        return "规律"
    actual_date = parse_date(actual_purchase)
    if pd.isna(actual_date) or pd.isna(expected_date):
        return "超期"
    if actual_date > expected_date or actual_date < expected_date - pd.Timedelta(days=5):
        return "超期"
    return "规律"


def reason_state(status: str, value, label: str) -> str:
    """Compute assessment label for 延期用药原因/未购药原因 columns."""
    text = clean_text(value)
    if status == "规律":
        return "" if not text else f"错误：规律但存在{label}"
    if status == "超期":
        return text if text else f"错误：超期但无{label}"
    return ""


def rule_youhede_sale(sale_row: pd.Series, follow_row: pd.Series, history: pd.DataFrame = None) -> tuple[str, str, str]:
    """优赫得 rule: returns (issue_reason, delay_state_label, no_purchase_state_label).
    Uses sales history for purchase timing check if provided."""
    expected = parse_date(sale_row.get("销售时间")) + pd.Timedelta(days=21)
    actual = transaction_actual_purchase(expected, follow_row, history)
    status = purchase_status_label(actual, expected)
    no_purchase_state = reason_state(status, follow_row.get("本月未购药的原因"), "未购药原因")
    delay_state = reason_state(status, follow_row.get("患者延迟用药的原因"), "延期用药原因")

    # If expected date is in the future and purchase hasn't happened yet
    if pd.notna(expected) and expected > pd.Timestamp.today().normalize() and clean_text(actual) == "日期未到":
        return (f"预计{expected.strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因", delay_state, no_purchase_state)

    # No follow-up done
    if pd.isna(parse_date(follow_row.get("执行时间"))):
        if status == "超期":
            return (f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}，目前已经超期，已超期但未记录未购药原因或延期用药原因，需补充", delay_state, no_purchase_state)
        return (f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}", delay_state, no_purchase_state)

    # Overdue with no purchase near expected date (actual = NaT or too early/late)
    if status == "超期":
        # Genuinely overdue: no purchase near expected date
        if pd.isna(parse_date(actual)):
            # Patient hasn't purchased near the expected date
            has_delay_error = not delay_state or "错误" in delay_state
            has_no_purchase_error = not no_purchase_state or "错误" in no_purchase_state
            if has_delay_error and has_no_purchase_error:
                return ("已超期但未记录未购药原因或延期用药原因，需补充", delay_state, no_purchase_state)
            if has_delay_error:
                return ("患者延迟用药，应填写延迟用药原因", delay_state, no_purchase_state)
            if has_no_purchase_error:
                return ("已超期但未记录未购药原因，需补充", delay_state, no_purchase_state)
            # Has both reasons but something might still be wrong
            return ("", delay_state, no_purchase_state)

        # Overdue but purchased late (actual date > expected)
        actual_date = parse_date(actual)
        if pd.notna(actual_date) and pd.notna(expected):
            if actual_date > expected and (not delay_state or "错误" in delay_state):
                return ("患者延迟用药，应填写延迟用药原因", delay_state, no_purchase_state)
            if (not delay_state or "错误" in delay_state) and (not no_purchase_state or "错误" in no_purchase_state):
                return ("已超期但未记录未购药原因或延期用药原因，需补充", delay_state, no_purchase_state)

    # On-time purchase but has no_purchase_reason → need to verify
    if status == "规律" and no_purchase_state:
        if "错误" in no_purchase_state:
            return ("需填写本月未购药原因/延期用药原因", delay_state, no_purchase_state)
        return ("购药规律但记录了未购药原因，需核实", delay_state, no_purchase_state)

    return ("", delay_state, no_purchase_state)


def rule_fanshuzhuo_sale(sale_row: pd.Series, follow_row: pd.Series, history: pd.DataFrame = None) -> tuple[str, str, str]:
    """凡舒卓 rule: returns (issue_reason, delay_state_label, no_purchase_state_label).
    Uses sales history for purchase timing check if provided."""
    expected = parse_date(sale_row.get("销售时间")) + pd.Timedelta(days=28)
    actual = transaction_actual_purchase(expected, follow_row, history)
    status = purchase_status_label(actual, expected)
    no_purchase_state = reason_state(status, follow_row.get("本月未购药的原因"), "未购药原因")

    # No follow-up done
    if pd.isna(parse_date(follow_row.get("执行时间"))):
        if status == "超期":
            return (f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}，目前已经超期，已超期但未记录本月未购药原因，需补充", "", no_purchase_state)
        return (f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}", "", no_purchase_state)

    # Expected date in the future
    if pd.notna(expected) and expected > pd.Timestamp.today().normalize() and clean_text(actual) == "日期未到":
        return (f"预计{expected.strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因", "", no_purchase_state)

    # Overdue: no_purchase_reason missing
    if status == "超期" and (not no_purchase_state or "错误" in no_purchase_state):
        return ("已超期但未记录本月未购药原因，需补充", "", no_purchase_state)

    # On-time purchase but has no_purchase_reason → verify
    if status == "规律" and no_purchase_state:
        if "错误" in no_purchase_state:
            return ("需填写本月未购药原因", "", no_purchase_state)
        return ("购药规律但记录了未购药原因，需核实", "", no_purchase_state)

    return ("", "", no_purchase_state)


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
    # Use follow-up data only for stock determination, dedup per patient+pharmacy
    # Prefer report file rows over detail rows for overlapping patients
    # (利普卓暂用旧逻辑，待后续单独矫正)
    lp_follow = follow_df[follow_df["__brand"].eq("利普卓")].copy() if not follow_df.empty else pd.DataFrame()
    if not lp_follow.empty:
        lp_follow_sorted = lp_follow.sort_values(
            ["__is_report", "__date"], ascending=[False, False]
        )
        lp_follow_dedup = lp_follow_sorted.drop_duplicates(subset=["__name", "__store"])
        for _, follow_row in lp_follow_dedup.iterrows():
            if not any([
                clean_text(follow_row.get("患者姓名")),
                clean_text(follow_row.get("患者手机号")),
                clean_text(follow_row.get("会员号")),
            ]):
                continue
            reason = rule_stock_brand_follow(follow_row, target_month)
            if not reason:
                continue
            rows.append(build_output_row("利普卓", pd.Series(dtype=object), follow_row, reason))

    # ─── 优赫得 ───
    # Per-patient approach: for each patient, use their most recent sale record
    # to determine expected_date, then check if they have follow-up issues
    yh_sales = sales_df[
        sales_df["__brand"].eq("优赫得")
        & sales_df["__store"].isin(["南充药房", "德阳关爱药房"])
    ].copy()
    yh_follow = follow_df[follow_df["__brand"].eq("优赫得")].copy()
    yh_processed = set()
    for key, history in yh_sales.groupby(yh_sales.apply(sale_patient_key, axis=1)):
        if not key:
            continue
        # Dedup: each patient appears only once
        if key in yh_processed:
            continue
        yh_processed.add(key)

        # Use the LAST sale record to determine expected_date
        last_sale = history.sort_values("__sale_date").iloc[-1]
        sale_date = parse_date(last_sale.get("销售时间"))
        if pd.isna(sale_date):
            continue
        expected = sale_date + pd.Timedelta(days=21)

        # Include patients whose expected_date is near the target month:
        # May-July expected dates cover overdue, target-month, and near-future
        if expected.to_period("M") > target_period(target_month) + 1:
            continue
        if expected.to_period("M") < target_period(target_month) - 1:
            continue

        # Find matching follow-up record
        follow_rows = yh_follow[yh_follow["__key_name_store"].eq(key)].copy()
        if not follow_rows.empty:
            follow_row = follow_rows.sort_values("__date", ascending=False).iloc[0]
        else:
            follow_row = find_follow_by_sale(yh_follow, last_sale)

        issue_reason, delay_label, no_purchase_label = rule_youhede_sale(last_sale, follow_row if not follow_row.empty else pd.Series(dtype=object), history)
        if issue_reason:
            rows.append(build_output_row("优赫得", last_sale, follow_row, issue_reason, delay_label, no_purchase_label))

    # ─── 凡舒卓 ───
    # Per-patient approach: for each patient, use their most recent sale record
    # to determine expected_date, then check if they have follow-up issues
    fsz_sales = sales_df[sales_df["__brand"].eq("凡舒卓")].copy()
    fsz_follow = follow_df[follow_df["__brand"].eq("凡舒卓")].copy()
    fsz_processed = set()
    for key, history in fsz_sales.groupby(fsz_sales.apply(sale_patient_key, axis=1)):
        if not key:
            continue
        # Dedup: each patient appears only once
        if key in fsz_processed:
            continue
        fsz_processed.add(key)

        # Use the LAST sale record to determine expected_date
        last_sale = history.sort_values("__sale_date").iloc[-1]
        sale_date = parse_date(last_sale.get("销售时间"))
        if pd.isna(sale_date):
            continue
        expected = sale_date + pd.Timedelta(days=28)

        # Include patients whose expected_date is in or before target month + 1 month
        # (not yet due but approaching), or overdue from recent months
        if expected.to_period("M") > target_period(target_month) + 1:
            continue
        if expected.to_period("M") < target_period(target_month) - 5:
            continue

        # Find matching follow-up record
        follow_rows = fsz_follow[fsz_follow["__key_name_store"].eq(key)].copy()
        if not follow_rows.empty:
            follow_row = follow_rows.sort_values("__date", ascending=False).iloc[0]
        else:
            follow_row = find_follow_by_sale(fsz_follow, last_sale)

        issue_reason, delay_label, no_purchase_label = rule_fanshuzhuo_sale(last_sale, follow_row if not follow_row.empty else pd.Series(dtype=object), history)
        if issue_reason:
            rows.append(build_output_row("凡舒卓", last_sale, follow_row, issue_reason, delay_label, no_purchase_label))

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
