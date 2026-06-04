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
    "来源文件",
    "随访表类型",
]

BRANDS = ["泰瑞沙", "利普卓", "英飞凡", "荃科得", "优赫得", "凡舒卓"]
SPECIAL_FOLLOW_BRANDS = {"优赫得", "荃科得"}


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


def classify_follow_file(frame: pd.DataFrame, filename: str) -> str:
    if any(brand in filename for brand in SPECIAL_FOLLOW_BRANDS):
        return "优赫得/荃科得专属表"
    if any(brand in filename for brand in set(BRANDS) - SPECIAL_FOLLOW_BRANDS):
        return "通用随访表"

    columns = set(frame.columns)
    has_doctor_duration = bool({"医生建议用药时长", "医生建议服用时间"} & columns)
    has_month_purchase = "当月是否购药" in columns
    if has_month_purchase and not has_doctor_duration:
        return "优赫得/荃科得专属表"
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
        prepared.append((file_type, frame))
    file_types = {file_type for file_type, _ in prepared}
    should_route_by_brand = len(prepared) > 1 and "优赫得/荃科得专属表" in file_types

    frames = []
    for file_type, frame in prepared:
        if should_route_by_brand:
            if file_type == "优赫得/荃科得专属表":
                frame = frame[frame["__brand"].isin(SPECIAL_FOLLOW_BRANDS)].copy()
            else:
                frame = frame[~frame["__brand"].isin(SPECIAL_FOLLOW_BRANDS)].copy()
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


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


def find_sale_by_name_store(row: pd.Series, lookup: dict[str, dict]) -> dict:
    key = clean_text(row.get("__key_name_store"))
    if key in lookup:
        return lookup[key]
    if "|" in key:
        any_key = "ANY|" + key.split("|", 1)[1]
        if any_key in lookup:
            return lookup[any_key]
    return find_sale(row, lookup)


def yes_value(value) -> bool | None:
    text = clean_text(value)
    if text in {"是", "已购药", "有", "Y", "yes", "Yes"}:
        return True
    if text in {"否", "未购药", "无", "N", "no", "No"}:
        return False
    return None


def should_have_no_purchase_reason(sale: dict, target_month: str) -> str:
    monthly_qty = sale.get("monthly_qty", {})
    if not monthly_qty:
        return "否" if sale.get("purchased_this_month") else "是"

    target = pd.Period(target_month, freq="M")
    if monthly_qty.get(target, 0) > 0:
        return "否"
    for months_back in range(1, 22):
        total = sum(monthly_qty.get(target - offset, 0) for offset in range(0, months_back + 1))
        if total > months_back:
            return "否"
    return "是"


def purchased_in_target_month(row: pd.Series, target_month: str) -> bool | None:
    follow_purchase = yes_value(row.get("当月是否购药"))
    if follow_purchase is not None:
        return follow_purchase

    last_purchase = parse_date(row.get("会员最近一次门店购药时间"))
    if pd.notna(last_purchase):
        return last_purchase.to_period("M") == pd.Period(target_month, freq="M")
    return None


def template_purchase_status(row: pd.Series, sale: dict, target_month: str) -> bool | None:
    inferred_purchase = purchased_in_target_month(row, target_month)
    return inferred_purchase if inferred_purchase is not None else sale.get("purchased_this_month")


def rule_common_sales_matrix(row: pd.Series, sale: dict, target_month: str) -> str:
    indication = clean_text(row.get("适应症"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    should_have_reason = sale.get("should_have_no_purchase_reason", "是")

    reason = ""
    if not indication:
        reason += "适应症未填写需要填写适应症"
    if not indication and ((should_have_reason == "是" and not no_purchase_reason) or (should_have_reason == "否" and no_purchase_reason)):
        reason += ", "
    if should_have_reason == "是" and not no_purchase_reason:
        reason += "患者本月还未来购药且没有存药，需填写本月未购药原因"
    if should_have_reason == "否" and no_purchase_reason:
        reason += "患者本月已来购药，本月未购药原因应该为空"
    return reason.strip()


def rule_yingfeifan(row: pd.Series, sale: dict, target_month: str) -> str:
    indication = clean_text(row.get("适应症"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    follow_purchase = yes_value(row.get("当月是否购药"))
    purchase_status = bool(sale.get("purchased_this_month", False))

    purchase_reason = ""
    if purchase_status is True:
        if not (follow_purchase is True and not no_purchase_reason):
            purchase_reason = "本月患者已购药，请检查当月是否购药列和本月未购药原因是否准确"
    else:
        if not (follow_purchase is False and no_purchase_reason):
            purchase_reason = "本月患者未购药，请检查当月是否购药列和本月未购药原因是否准确"

    indication_reason = ""
    if not indication:
        indication_reason = "适应症为空需重新生成随访进行填写"

    if purchase_reason and indication_reason:
        return purchase_reason + "；" + indication_reason
    return purchase_reason or indication_reason


def expected_next_purchase(row: pd.Series, sale: dict):
    recent_date = parse_date(row.get("会员最近一次门店购药时间"))
    if pd.isna(recent_date):
        recent_date = sale.get("last_sale_date", pd.NaT)
    if pd.isna(recent_date):
        return pd.NaT
    brand = clean_text(row.get("__brand"))
    days = 21 if brand in {"优赫得", "荃科得"} else 28
    return recent_date + pd.Timedelta(days=days)


def rule_transaction_follow(row: pd.Series, sale: dict, target_month: str) -> str:
    reasons = []
    brand = clean_text(row.get("__brand"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    delay_reason = clean_text(row.get("患者延迟用药的原因"))
    purchase_status = template_purchase_status(row, sale, target_month)
    next_date = expected_next_purchase(row, sale)

    if purchase_status is None:
        if pd.notna(next_date):
            return f"应做随访未做随访，该患者预计购药日期为：{next_date.strftime('%Y-%m-%d')}"
        return "应做随访未做随访"

    if pd.notna(next_date) and next_date > pd.Timestamp.today():
        return f"预计{next_date.strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因"

    if purchase_status is False:
        if brand in {"优赫得", "凡舒卓"} and not delay_reason:
            reasons.append("已超期但未记录延期用药原因，需补充")
        if not no_purchase_reason:
            reasons.append("已超期但未记录本月未购药原因，需补充")

    if purchase_status is True and no_purchase_reason:
        reasons.append("需填写本月未购药原因/延期用药原因")

    return "；".join(reasons)


def transaction_actual_purchase(expected_date, follow_row: pd.Series):
    follow_recent = parse_date(follow_row.get("会员最近一次门店购药时间"))
    if pd.isna(expected_date):
        return pd.NaT
    today = pd.Timestamp.today().normalize()
    if expected_date <= today:
        return follow_recent
    if pd.notna(follow_recent) and follow_recent >= expected_date - pd.Timedelta(days=5):
        return follow_recent
    return "日期未到"


def purchase_status(actual_purchase, expected_date) -> str:
    if clean_text(actual_purchase) == "日期未到":
        return "规律"
    actual_date = parse_date(actual_purchase)
    if pd.isna(actual_date) or pd.isna(expected_date):
        return ""
    if actual_date > expected_date or actual_date < expected_date - pd.Timedelta(days=5):
        return "超期"
    return "规律"


def reason_state(status: str, value, label: str) -> str:
    text = clean_text(value)
    if status == "规律":
        return "" if not text else f"错误：规律但存在{label}"
    if status == "超期":
        return text if text else f"错误：超期但无{label}"
    return ""


def rule_youhede_sale(sale_row: pd.Series, follow_row: pd.Series) -> str:
    expected = parse_date(sale_row.get("销售时间")) + pd.Timedelta(days=21)
    actual = transaction_actual_purchase(expected, follow_row)
    status = purchase_status(actual, expected)
    no_purchase_state = reason_state(status, follow_row.get("本月未购药的原因"), "未购药原因")
    delay_state = reason_state(status, follow_row.get("患者延迟用药的原因"), "延期用药原因")

    if pd.notna(expected) and expected > pd.Timestamp.today().normalize() and clean_text(actual) == "日期未到":
        return f"预计{expected.strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因"
    if status == "超期" and ("错误" in no_purchase_state or not no_purchase_state) and ("错误" in delay_state or not delay_state):
        return "已超期但未记录延期用药原因，需补充"
    actual_date = parse_date(actual)
    if pd.notna(actual_date) and pd.notna(expected) and actual_date > expected and actual_date.month == expected.month and actual_date.year == expected.year:
        return "" if delay_state and "错误" not in delay_state else "患者延迟用药，应填写延迟用药原因"
    if no_purchase_state == "错误：规律但存在未购药原因":
        return "需填写本月未购药原因/延期用药原因"
    if no_purchase_state == "错误：超期但无未购药原因":
        return "待观察，如果本月内未来购药，需重新生成随访填写本月未购药原因"
    if status == "规律" and no_purchase_state:
        return "购药规律但记录了未购药原因，需核实"
    return ""


def rule_fanshuzhuo_sale(sale_row: pd.Series, follow_row: pd.Series) -> str:
    expected = parse_date(sale_row.get("销售时间")) + pd.Timedelta(days=28)
    actual = transaction_actual_purchase(expected, follow_row)
    status = purchase_status(actual, expected)
    no_purchase_state = reason_state(status, follow_row.get("本月未购药的原因"), "未购药原因")

    if pd.isna(parse_date(follow_row.get("执行时间"))):
        extra = "，目前已经超期，已超期但未记录本月未购药原因，需补充" if status == "超期" else ""
        return f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}{extra}"
    if pd.notna(expected) and expected > pd.Timestamp.today().normalize() and clean_text(actual) == "日期未到":
        return f"预计{expected.strftime('%Y-%m-%d')}购药，如果未购药需填写延期用药原因"
    if status == "超期" and ("错误" in no_purchase_state or not no_purchase_state):
        return "已超期但未记录本月未购药原因，需补充"
    if no_purchase_state == "错误：规律但存在未购药原因":
        return "需填写本月未购药原因"
    if no_purchase_state == "错误：超期但无未购药原因":
        return "待观察，如果本月内未来购药，需重新生成随访填写本月未购药原因"
    if status == "规律" and no_purchase_state:
        return "购药规律但记录了未购药原因，需核实"
    return ""


def rule_quankede_dates(last_follow, second_follow) -> str:
    if pd.isna(last_follow):
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    if pd.isna(second_follow):
        return f"本月未完成二次随访，请在 {(last_follow + pd.Timedelta(days=10)).strftime('%Y-%m-%d')} 完成二次随访"
    days = abs((last_follow - second_follow).days)
    if days < 10:
        return f"本月已随访两次，但两次间隔＜十天，目前间隔 {days} 天，请重新生成随访"
    return ""


def rule_quankede(row: pd.Series, sale: dict, target_month: str) -> str:
    last_follow = parse_date(row.get("执行时间"))
    second_follow = parse_date(row.get("倒数第二次门店购药时间"))
    return rule_quankede_dates(last_follow, second_follow)


def build_issue_reason(row: pd.Series, sale: dict, target_month: str) -> str:
    brand = clean_text(row.get("__brand")) or clean_text(row.get("品牌"))
    if brand in {"泰瑞沙", "利普卓"}:
        return rule_common_sales_matrix(row, sale, target_month)
    if brand == "英飞凡":
        return rule_yingfeifan(row, sale, target_month)
    if brand in {"优赫得", "凡舒卓"}:
        return rule_transaction_follow(row, sale, target_month)
    if brand == "荃科得":
        return rule_quankede(row, sale, target_month)
    return ""


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


def build_output_row(brand: str, source_row: pd.Series, follow_row: pd.Series, reason: str) -> dict:
    row = follow_row if not follow_row.empty else source_row
    return {
        "品牌": brand,
        "药房": clean_text(row.get("门店")) or clean_text(source_row.get("药房名称")),
        "执行时间": row.get("执行时间", ""),
        "患者姓名": clean_text(row.get("患者姓名")) or clean_text(source_row.get("会员姓名")) or clean_text(source_row.get("开票抬头")),
        "联系方式": clean_text(row.get("患者手机号")) or clean_text(source_row.get("会员电话")),
        "适应症": clean_text(row.get("适应症")) or clean_text(source_row.get("适应症")),
        "会员最近一次门店购药时间": row.get("会员最近一次门店购药时间", source_row.get("销售时间", "")),
        "会员最近一次门店购药盒数": row.get("会员最近一次门店购药盒数", source_row.get("销售数量", "")),
        "延期用药原因": clean_text(row.get("患者延迟用药的原因")),
        "本月未购药的原因": clean_text(row.get("本月未购药的原因")),
        "医生建议用药时长": clean_text(row.get("医生建议用药时长")),
        "需要补充的原因": reason,
        "来源文件": clean_text(row.get("来源文件")),
        "随访表类型": clean_text(row.get("随访表类型")),
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


def sales_for_key(sales_df: pd.DataFrame, key: str) -> pd.DataFrame:
    if not key or sales_df.empty:
        return pd.DataFrame()
    return sales_df[sales_df.apply(sale_patient_key, axis=1).eq(key)].copy()


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


def rule_stock_brand_follow(row: pd.Series, sale: dict, target_month: str) -> str:
    indication = clean_text(row.get("适应症"))
    no_purchase_reason = clean_text(row.get("本月未购药的原因"))
    should_have_reason = sale.get("should_have_no_purchase_reason", "是")
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
        return "本月未完成随访，请补充随访"

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


def transaction_reason(brand: str, sale_row: pd.Series, follow_row: pd.Series, history: pd.DataFrame, target_month: str, cycle_days: int) -> str:
    sale_date = parse_date(sale_row.get("销售时间"))
    if pd.isna(sale_date):
        return ""
    expected = sale_date + pd.Timedelta(days=cycle_days)
    if expected.to_period("M") != target_period(target_month):
        return ""

    if follow_row.empty:
        return f"应做随访未做随访，该患者预计购药日期为：{expected.strftime('%Y-%m-%d')}"

    no_purchase_reason = clean_text(follow_row.get("本月未购药的原因"))
    delay_reason = clean_text(follow_row.get("患者延迟用药的原因"))
    on_time = has_on_time_purchase(history, expected)
    later_purchase = has_later_target_purchase(history, expected, target_month)
    month_purchase = has_target_purchase(history, target_month)

    if on_time:
        if no_purchase_reason:
            return "购药规律但记录了未购药原因，需核实"
        return ""

    if brand == "优赫得":
        if later_purchase:
            return "" if delay_reason else "患者延迟用药，应填写延迟用药原因"
        if not month_purchase:
            return "" if no_purchase_reason else "已超期但未记录未购药原因或延期用药原因，需补充"
        return ""

    if brand == "凡舒卓":
        return "" if no_purchase_reason else "已超期但未记录本月未购药原因，需补充"

    return ""


def quankede_reason_for_patient(follow_rows: pd.DataFrame, last_sale_row: pd.Series, target_month: str) -> str:
    if follow_rows.empty or "__date" not in follow_rows.columns:
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    target_rows = follow_rows[follow_rows["__date"].map(lambda value: in_target_month(value, target_month))].copy()
    dates = target_rows["__date"].map(parse_date).dropna().sort_values(ascending=False)
    if len(dates) == 0:
        return "还未完成随访，请本月完成两次随访且间隔大于十天"
    if len(dates) == 1:
        return f"本月未完成二次随访，请在 {(dates.iloc[0] + pd.Timedelta(days=10)).strftime('%Y-%m-%d')} 完成二次随访"
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
    sales_lookup = build_sales_lookup(sales_df, target_month)

    rows = []
    for brand in ["优赫得", "凡舒卓", "荃科得"]:
        brand_sales = sales_df[sales_df["__brand"].eq(brand)].copy()
        if brand == "优赫得":
            brand_sales = brand_sales[brand_sales["药房名称"].map(clean_text).str.contains("南充|德阳关爱|泰山路", na=False)]
        if brand == "荃科得":
            brand_sales = brand_sales[brand_sales["__sale_date"].map(lambda value: in_target_month(value, target_month))]

        for key, history in brand_sales.groupby(brand_sales.apply(sale_patient_key, axis=1)):
            if not key:
                continue
            follow_brand = follow_df[follow_df["__brand"].eq(brand)]
            follow_rows = follow_brand[follow_brand["__key_name_store"].eq(key)].copy()
            if follow_rows.empty:
                sale_row = history.sort_values("__sale_date").iloc[-1]
                follow_row = find_follow_by_sale(follow_brand, sale_row)
            else:
                follow_row = follow_rows.sort_values("__date", ascending=False).iloc[0]

            if brand == "优赫得":
                candidate_sales = history[history["__sale_date"].map(lambda value: pd.notna(parse_date(value)) and (parse_date(value) + pd.Timedelta(days=21)).to_period("M") == target_period(target_month))]
                for _, sale_row in candidate_sales.iterrows():
                    reason = transaction_reason(brand, sale_row, follow_row, history, target_month, 21)
                    if reason:
                        rows.append(build_output_row(brand, sale_row, follow_row, reason))
            elif brand == "凡舒卓":
                candidate_sales = history[history["__sale_date"].map(lambda value: pd.notna(parse_date(value)) and (parse_date(value) + pd.Timedelta(days=28)).to_period("M") == target_period(target_month))]
                for _, sale_row in candidate_sales.iterrows():
                    reason = transaction_reason(brand, sale_row, follow_row, history, target_month, 28)
                    if reason:
                        rows.append(build_output_row(brand, sale_row, follow_row, reason))
            else:
                if history.empty:
                    continue
                sale_row = history.sort_values("__sale_date").iloc[-1]
                reason = quankede_reason_for_patient(follow_rows, sale_row, target_month)
                if reason:
                    rows.append(build_output_row(brand, sale_row, follow_row, reason))

    for _, follow_row in follow_df.iterrows():
        brand = clean_text(follow_row.get("__brand")) or clean_text(follow_row.get("品牌"))
        if brand not in {"泰瑞沙", "利普卓"}:
            continue

        if not any(
            [
                clean_text(follow_row.get("患者姓名")),
                clean_text(follow_row.get("患者手机号")),
                clean_text(follow_row.get("会员号")),
            ]
        ):
            continue

        sale = find_sale_by_name_store(follow_row, sales_lookup)
        reason = rule_stock_brand_follow(follow_row, sale, target_month)
        if not reason:
            continue

        rows.append(build_output_row(brand, pd.Series(dtype=object), follow_row, reason))

    yingfei_sales = sales_df[
        sales_df["__brand"].eq("英飞凡")
        & sales_df["药房名称"].map(clean_text).str.contains("晟德", na=False)
        & sales_df["__sale_date"].map(lambda value: pd.notna(parse_date(value)) and parse_date(value) >= pd.Timestamp("2025-01-01"))
    ].copy()
    yingfei_follow = follow_df[follow_df["__brand"].eq("英飞凡")].copy()
    for key, history in yingfei_sales.groupby(yingfei_sales.apply(person_key, axis=1)):
        if not key:
            continue
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
        follow_df.head(5000).to_excel(writer, index=False, sheet_name="随访底表预览")
        sales_df.head(5000).to_excel(writer, index=False, sheet_name="销售底表预览")
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
    st.write("如果上传多份随访底表，优赫得、荃科得会自动走专属表，其他品种走通用表。")
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
