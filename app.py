import streamlit as st
import pandas as pd
import re
import io
import pyreadstat

st.set_page_config(page_title="學校問卷回應整理工具（SPSS）", layout="wide")
st.title("學校問卷回應整理工具（SPSS .sav）")
st.markdown("""
上傳 SPSS `.sav` 檔案，自動產生四張表：
- SchName x Q1Q2_12（由 Q2_12 + Q1_12 合併）
- SchName x Q1Q2_13
- SchName x Q1Q2_14
- SchName x Q1Q2_15

每張表皆：
- 每條老師回應一列
- 學校欄只第一次顯示
- 可分別下載 CSV
""")

# ---------------------------
# 輔助函數：學校欄只第一次顯示
# ---------------------------

def process_school_column(df, school_col="SchName", text_col="text"):
    """
    學校欄：同一學校只第一次顯示，後續留空
    """
    df = df.copy()
    df["學校顯示"] = df[school_col].where(df[school_col].shift() != df[school_col], "")
    out = df[["學校顯示", text_col]].rename(columns={"學校顯示": "SchName", text_col: "text"})
    return out


# ---------------------------
# 上傳 .sav 檔案
# ---------------------------

uploaded = st.file_uploader(
    "上傳 SPSS .sav 檔案",
    type=["sav"],
)

if uploaded is not None:
    try:
        # 使用 pyreadstat 讀取 SPSS .sav 檔案
        _, raw_df = pyreadstat.read_sav(uploaded)

        st.subheader("原始資料預覽")
        st.dataframe(raw_df.head(20), use_container_width=True)
        st.write("欄位列表：", list(raw_df.columns))

        # 確認必要欄位是否存在
        required_cols = ["SchName", "Q2_12", "Q1_12", "Q1Q2_13", "Q1Q2_14", "Q1Q2_15"]
        missing = [c for c in required_cols if c not in raw_df.columns]
        if missing:
            st.error(f"缺少必要欄位：{missing}")
            st.stop()

        # ---------------------------
        # 處理 Q1Q2_12 = Q2_12 + Q1_12
        # ---------------------------

        df = raw_df[["SchName", "Q2_12", "Q1_12", "Q1Q2_13", "Q1Q2_14", "Q1Q2_15"]].copy()

        # 合併 Q2_12 和 Q1_12 為 Q1Q2_12
        def merge_q12(row):
            parts = []
            if pd.notna(row["Q2_12"]) and str(row["Q2_12"]).strip() != "":
                parts.append(str(row["Q2_12"]).strip())
            if pd.notna(row["Q1_12"]) and str(row["Q1_12"]).strip() != "":
                parts.append(str(row["Q1_12"]).strip())
            return " ".join(parts)

        df["Q1Q2_12"] = df.apply(merge_q12, axis=1)

        # 刪去 Q1Q2_12 為空 / nil / / 的行（只針對 Q1Q2_12 表）
        def is_valid(q):
            if pd.isna(q):
                return False
            qs = str(q).strip()
            if qs == "":
                return False
            if qs.lower() in ["nil", "/", "nan"]:
                return False
            return True

        # ---------------------------
        # 產生四張表
        # ---------------------------

        tables = {}
        text_cols = {
            "Q1Q2_12": "Q1Q2_12",
            "Q1Q2_13": "Q1Q2_13",
            "Q1Q2_14": "Q1Q2_14",
            "Q1Q2_15": "Q1Q2_15",
        }

        for key, col in text_cols.items():
            df_tmp = df[["SchName", col]].copy()
            if key == "Q1Q2_12":
                # 只保留 Q1Q2_12 有效的行
                df_tmp = df_tmp[df_tmp[col].apply(is_valid)].reset_index(drop=True)
            else:
                # 其他表也建議過濾空白
                df_tmp = df_tmp[df_tmp[col].apply(is_valid)].reset_index(drop=True)

            df_out = process_school_column(df_tmp, school_col="SchName", text_col=col)
            df_out = df_out.rename(columns={"text": col})
            tables[key] = df_out

        # ---------------------------
        # 顯示和下載
        # ---------------------------

        st.subheader("整理後表格")

        for key, df_out in tables.items():
            st.markdown(f"### {key}")
            st.dataframe(df_out.head(50), use_container_width=True)

            buffer = io.BytesIO()
            df_out.to_csv(buffer, index=False, encoding="utf-8-sig")
            buffer.seek(0)

            st.download_button(
                label=f"下載 {key}.csv",
                data=buffer,
                file_name=f"{key}.csv",
                mime="text/csv",
                key=f"download_{key}",
            )

    except Exception as e:
        st.error(f"讀取或處理檔案時發生錯誤：{e}")
        st.info("請確認 .sav 檔案中包含 SchName, Q2_12, Q1_12, Q1Q2_13, Q1Q2_14, Q1Q2_15 等欄位。")

else:
    st.info("尚未上傳 .sav 檔案。")