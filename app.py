import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="學校問卷回應整理工具", layout="wide")
st.title("學校問卷回應整理工具")
st.markdown("上傳包含「學校、問題 1、問題 2」三列的資料（CSV / Excel / TXT），自動整理成每條老師回應一列、學校只第一次顯示的表格。")

# ---------------------------
# 輔助函數：從原始文字解析記錄
# ---------------------------

def parse_raw_text(text):
    """
    假設資料是三列一組：[學校, 問題1, 問題2]，但實際可能有很多只有學校或只有問題的行。
    規則：
    - 以 P+數字 開頭且長度較長的行視為學校行
    - 學校行之後的非學校行視為該學校的問題內容（最多兩行：問題1、問題2）
    - 空白、nil、/ 視為無內容
    """
    lines = [l for l in text.splitlines()]

    # 自動偵測學校前綴（例如 P03, P07, P08, P09, P10, P32, P33 等）
    # 找所有以 P+數字 開頭的行，提取前綴
    prefixes = set()
    for l in lines:
        m = re.match(r"^(P\d+)", l.strip())
        if m:
            prefixes.add(m.group(1))
    prefixes = sorted(prefixes, key=lambda x: len(x), reverse=True)  # 較長前綴優先

    if not prefixes:
        # 若完全找不到 P 開頭，勉強用前幾個字當學校
        def is_school_line(line):
            return len(line) > 10
    else:
        def is_school_line(line):
            if len(line) < 6:
                return False
            for p in prefixes:
                if line.startswith(p):
                    return True
            return False

    records = []
    current_school = None
    current_q1 = None
    current_q2 = None

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line == "":
            i += 1
            continue

        if is_school_line(line):
            # 保存上一組
            if current_school is not None:
                q_parts = [x for x in [current_q1, current_q2] if x and x.strip() != ""]
                if q_parts:
                    q = " ".join(q_parts)
                    if q.strip().lower() not in ["nil", "/", ""]:
                        records.append({"學校": current_school, "合併問題": q.strip()})

            # 新學校
            current_school = line
            current_q1 = None
            current_q2 = None
            i += 1

            # 下一行非學校行 -> q1
            if i < len(lines):
                nxt = lines[i].strip()
                if not is_school_line(nxt) and nxt != "":
                    current_q1 = nxt
                    i += 1

            # 再下一行非學校行 -> q2
            if i < len(lines):
                nxt = lines[i].strip()
                if not is_school_line(nxt) and nxt != "":
                    current_q2 = nxt
                    i += 1

            # 清理 nil / /
            if current_q1 and current_q1.lower() in ["nil", "/", ""]:
                current_q1 = None
            if current_q2 and current_q2.lower() in ["nil", "/", ""]:
                current_q2 = None
        else:
            i += 1

    # 最後一組
    if current_school is not None:
        q_parts = [x for x in [current_q1, current_q2] if x and x.strip() != ""]
        if q_parts:
            q = " ".join(q_parts)
            if q.strip().lower() not in ["nil", "/", ""]:
                records.append({"學校": current_school, "合併問題": q.strip()})

    return pd.DataFrame(records)


def process_school_column(df):
    """
    學校欄：同一學校只第一次顯示，後續留空
    假設 df 已有 "學校" 和 "合併問題" 兩欄
    """
    df = df.copy()
    df["學校顯示"] = df["學校"].where(df["學校"].shift() != df["學校"], "")
    out = df[["學校顯示", "合併問題"]].rename(columns={"學校顯示": "學校"})
    return out


# ---------------------------
# 上傳檔案
# ---------------------------

uploaded = st.file_uploader(
    "上傳資料檔案（CSV / Excel / TXT）",
    type=["csv", "xlsx", "xls", "txt"],
)

if uploaded is not None:
    try:
        if uploaded.name.endswith(".csv"):
            raw_df = pd.read_csv(uploaded)
        elif uploaded.name.endswith(".txt"):
            text = uploaded.read().decode("utf-8")
            raw_df = parse_raw_text(text)
        else:
            # Excel
            raw_df = pd.read_excel(uploaded)

        st.subheader("原始資料預覽")
        st.dataframe(raw_df.head(20), use_container_width=True)

        # 判斷欄位
        cols = list(raw_df.columns)
        st.write("偵測到的欄位：", cols)

        # 自動匹配學校、問題1、問題2
        school_col = None
        q1_col = None
        q2_col = None

        # 簡單規則：包含「學校」或第一欄當學校；其餘當問題
        for c in cols:
            cl = str(c).lower()
            if "學校" in cl or "school" in cl:
                school_col = c
                break
        if school_col is None and len(cols) >= 1:
            school_col = cols[0]

        remaining = [c for c in cols if c != school_col]
        if len(remaining) >= 1:
            q1_col = remaining[0]
        if len(remaining) >= 2:
            q2_col = remaining[1]

        st.write(f"使用欄位：學校={school_col}, 問題1={q1_col}, 問題2={q2_col}")

        if school_col and q1_col:
            df = raw_df[[c for c in cols if c in [school_col, q1_col, q2_col]]].copy()
            df = df.rename(columns={school_col: "學校", q1_col: "問題1", q2_col: "問題2" if q2_col else "問題2"})
            if "問題2" not in df.columns:
                df["問題2"] = ""

            # 合併問題
            def merge_q(row):
                parts = []
                if pd.notna(row["問題1"]) and str(row["問題1"]).strip() != "":
                    parts.append(str(row["問題1"]).strip())
                if pd.notna(row["問題2"]) and str(row["問題2"]).strip() != "":
                    parts.append(str(row["問題2"]).strip())
                return " ".join(parts)

            df["合併問題"] = df.apply(merge_q, axis=1)
            df = df[["學校", "合併問題"]]

            # 刪去空白 / nil / /
            def is_valid(q):
                if pd.isna(q):
                    return False
                qs = str(q).strip()
                if qs == "":
                    return False
                if qs.lower() in ["nil", "/", "nan"]:
                    return False
                return True

            df = df[df["合併問題"].apply(is_valid)].reset_index(drop=True)

            # 學校只第一次顯示
            df_out = process_school_column(df)

            st.subheader("整理後預覽")
            st.dataframe(df_out.head(50), use_container_width=True)

            # 下載
            buffer = io.BytesIO()
            df_out.to_csv(buffer, index=False, encoding="utf-8-sig")
            buffer.seek(0)

            st.download_button(
                label="下載整理後的 CSV",
                data=buffer,
                file_name="school_feedback_cleaned.csv",
                mime="text/csv",
            )
        else:
            st.warning("未能自動匹配到學校和問題欄位，請檢查檔案結構。")

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
        st.info("如果檔案是純文字三列格式（學校、問題 1、問題 2 交替），請改用 .txt 上傳。")

else:
    st.info("尚未上傳檔案。你可以先上傳一個包含學校和問題回應的 CSV / Excel / TXT 檔案。")

# ---------------------------
# 純文字模式（直接貼上）
# ---------------------------

st.subheader("或直接貼上純文字資料")
st.markdown("格式範例：每三行為一組【學校、問題 1、問題 2】，可有多組。")

raw_text = st.text_area(
    "貼上純文字資料",
    height=200,
    placeholder="P03某某小學\n回應 1\n回應 2\nP07另一小學\n回應 1\n回應 2\n...",
)

if raw_text.strip():
    try:
        df_raw = parse_raw_text(raw_text)
        if df_raw.empty:
            st.warning("未解析到任何有效記錄，請檢查格式。")
        else:
            df_out = process_school_column(df_raw)
            st.subheader("整理後預覽（純文字）")
            st.dataframe(df_out.head(50), use_container_width=True)

            buffer = io.BytesIO()
            df_out.to_csv(buffer, index=False, encoding="utf-8-sig")
            buffer.seek(0)

            st.download_button(
                label="下載整理後的 CSV（純文字）",
                data=buffer,
                file_name="school_feedback_cleaned_text.csv",
                mime="text/csv",
            )
    except Exception as e:
        st.error(f"解析純文字時發生錯誤：{e}")