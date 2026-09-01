import streamlit as st
import pandas as pd
import io
import os
import tempfile
import pyreadstat

st.set_page_config(page_title="學校問卷回應整理工具（SPSS）", layout="wide")
st.title("學校問卷回應整理工具（SPSS .sav）")
st.markdown("""
上傳 SPSS `.sav` 檔案後，工具會產生四張表：
- SchName x Q1Q2_12（由 Q2_12 和 Q1_12 合併）
- SchName x Q1Q2_13
- SchName x Q1Q2_14
- SchName x Q1Q2_15

每張表會保留每位老師的一條回應；同一學校只在首次出現時顯示名稱。
""")

INVALID_VALUES = {"", "nil", "/", "nan", "none", "n/a", "na"}
REQUIRED_COLUMNS = [
    "SchName", "Q2_12", "Q1_12", "Q1Q2_13", "Q1Q2_14", "Q1Q2_15"
]


def is_valid_response(value):
    """判斷回應是否有有效文字。"""
    if pd.isna(value):
        return False
    return str(value).strip().lower() not in INVALID_VALUES


def combine_responses(value1, value2):
    """合併兩個問題欄的有效文字。"""
    parts = []
    for value in (value1, value2):
        if is_valid_response(value):
            parts.append(str(value).strip())
    return " ".join(parts)


def make_question_table(data, question_column):
    """建立指定問題的兩欄表格，並讓連續重複的學校名稱只顯示一次。"""
    output = data[["SchName", question_column]].copy()
    output = output[output[question_column].apply(is_valid_response)].copy()
    output = output.reset_index(drop=True)

    output["學校"] = output["SchName"].where(
        output["SchName"].ne(output["SchName"].shift()), ""
    )
    output = output[["學校", question_column]]
    return output


def read_sav_file(uploaded_file):
    """把 Streamlit 上傳檔暫存後讀為 DataFrame。"""
    file_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as temp_file:
            temp_file.write(uploaded_file.getvalue())
            file_path = temp_file.name

        # pyreadstat 正確回傳順序是：(DataFrame, metadata)
        raw_df, metadata = pyreadstat.read_sav(file_path)
        return raw_df, metadata
    finally:
        if file_path and os.path.exists(file_path):
            os.remove(file_path)


uploaded = st.file_uploader("上傳 SPSS .sav 檔案", type=["sav"])

if uploaded is None:
    st.info("請上傳包含 SchName、Q2_12、Q1_12、Q1Q2_13、Q1Q2_14、Q1Q2_15 的 .sav 檔案。")
    st.stop()

try:
    raw_df, metadata = read_sav_file(uploaded)

    if not isinstance(raw_df, pd.DataFrame):
        st.error(f"無法把 .sav 檔讀取成資料表，取得的類型：{type(raw_df)}")
        st.stop()

    st.subheader("原始資料預覽")
    st.caption(f"共讀取 {len(raw_df):,} 筆資料、{len(raw_df.columns):,} 個變量。")
    st.dataframe(raw_df.head(20), use_container_width=True)

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in raw_df.columns]
    if missing_columns:
        st.error("檔案缺少必要變量：" + ", ".join(missing_columns))
        st.write("目前讀取到的變量：", list(raw_df.columns))
        st.stop()

    selected_columns = REQUIRED_COLUMNS.copy()
    data = raw_df[selected_columns].copy()

    # 合併問題 12：指定先放 Q2_12，後放 Q1_12。
    data["Q1Q2_12"] = data.apply(
        lambda row: combine_responses(row["Q2_12"], row["Q1_12"]), axis=1
    )

    question_columns = ["Q1Q2_12", "Q1Q2_13", "Q1Q2_14", "Q1Q2_15"]
    tables = {
        question: make_question_table(data, question)
        for question in question_columns
    }

    st.subheader("四張整理後表格")
    tabs = st.tabs(question_columns)

    for tab, question in zip(tabs, question_columns):
        with tab:
            result_df = tables[question]
            st.caption(f"有效回應：{len(result_df):,} 條")
            st.dataframe(result_df, use_container_width=True, height=500)

            csv_data = result_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                label=f"下載 {question}.csv",
                data=csv_data,
                file_name=f"SchName_x_{question}.csv",
                mime="text/csv",
                key=f"download_{question}",
            )

    # 一次下載四張表，方便直接用於報告製作。
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        for question, result_df in tables.items():
            result_df.to_excel(writer, index=False, sheet_name=question)
    excel_buffer.seek(0)

    st.download_button(
        label="一次下載四張表（Excel）",
        data=excel_buffer,
        file_name="SchName_feedback_tables.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_all_excel",
    )

except Exception as error:
    st.error(f"讀取或處理檔案時發生錯誤：{error}")
    st.exception(error)
