import io

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="DTP Excel 自动化平台",
    page_icon="📊",
    layout="wide",
)


st.title("DTP Excel 自动化平台")
st.caption("上传固定格式的 Excel 文件，自动完成清洗、计算、分析和导出。")


with st.sidebar:
    st.header("流程")
    st.markdown("1. 上传 Excel 模板")
    st.markdown("2. 系统自动识别字段")
    st.markdown("3. 查看计算结果")
    st.markdown("4. 下载处理后的 Excel")


uploaded_file = st.file_uploader(
    "上传 Excel 文件",
    type=["xlsx", "xls"],
    help="可以先上传销售底表、随访底表等固定格式文件。",
)


def read_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)


def to_excel_bytes(dataframe: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="结果")
    return output.getvalue()


if uploaded_file is None:
    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("上传模板")
        st.write("上传固定格式 Excel 文件，例如销售底表、随访底表或购药记录。")

    with col2:
        st.subheader("自动计算")
        st.write("后续可加入患者匹配、超期识别、二次随访判断和购药间隔分析。")

    with col3:
        st.subheader("输出结果")
        st.write("在线查看结果，并导出 Excel、随访任务或 BI 数据。")

    st.divider()
    st.subheader("可迁移的核心能力")
    features = [
        "XLOOKUP 患者匹配",
        "MAXIFS 末次随访判断",
        "购药时间间隔分析",
        "患者超期识别",
        "本月未购药原因提醒",
        "药房简称自动转换",
        "随访状态判断",
        "自动生成提醒话术",
    ]
    st.write(", ".join(features))
else:
    try:
        df = read_excel(uploaded_file)
    except Exception as exc:
        st.error(f"Excel 读取失败：{exc}")
        st.stop()

    st.success("文件读取成功")

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("行数", len(df))
    metric2.metric("列数", len(df.columns))
    metric3.metric("空值数量", int(df.isna().sum().sum()))

    st.subheader("数据预览")
    st.dataframe(df, use_container_width=True)

    st.subheader("字段列表")
    st.write(", ".join([str(column) for column in df.columns]))

    st.download_button(
        label="下载当前结果 Excel",
        data=to_excel_bytes(df),
        file_name="dtp_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
