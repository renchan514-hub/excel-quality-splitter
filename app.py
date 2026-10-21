# ============================================================
# Excel 資料品質分組 App
# 功能：
# 1. 上傳 .xlsx Excel 檔案
# 2. 讀取 Dataset_Combined 工作表
# 3. 嚴格按 Data_Quality_Label 分組
# 4. CLEAN -> 完整資料
# 5. DIRTY -> 不完整資料
# 6. 其他或空白 -> 無法識別資料
# 7. 產生可下載的 Excel 結果
# ============================================================

import io  # 處理記憶體中的 Excel 檔案。
import unicodedata  # 處理全形空格、零寬度空格等隱藏字元。
from pathlib import Path  # 處理輸出檔名。

import pandas as pd  # 讀取、處理及輸出 Excel。
import streamlit as st  # 建立 Streamlit 網頁介面。


# ============================================================
# Streamlit 頁面設定
# ============================================================

st.set_page_config(
    page_title="Excel 資料品質分組工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel 資料品質分組工具")
st.write(
    "上傳 Excel 檔案後，系統會根據原始檔的 "
    "`Data_Quality_Label` 欄位，將資料分為完整、不完整及無法識別三類。"
)

st.info(
    "分組規則：`CLEAN` 會進入「完整資料」；`DIRTY` 會進入「不完整資料」；"
    "空白、大小寫異常或其他標籤會進入「無法識別資料」。"
)


# ============================================================
# 可調整設定
# ============================================================

SOURCE_SHEET = "Dataset_Combined"  # 來源工作表名稱。
LABEL_COLUMN = "Data_Quality_Label"  # 原始檔的資料品質標籤欄位。
CLEAN_LABEL = "CLEAN"  # 完整資料標籤。
DIRTY_LABEL = "DIRTY"  # 不完整資料標籤。


# ============================================================
# 工具函式
# ============================================================

@st.cache_data
def clean_text(value):
    """
    將標籤文字標準化。

    可處理：
    - 前後空格：' CLEAN '
    - 連續空格：'CLEAN   '
    - 全形空格：'CLEAN　'
    - 大小寫：'clean'、'Clean'
    - 零寬度空格：'CLEAN\\u200b'
    """
    if pd.isna(value):  # 處理真正的空值。
        return ""

    text = str(value)  # 轉為文字。
    text = unicodedata.normalize("NFKC", text)  # 正規化全形空格等相似字元。
    text = text.replace("\u200b", "")  # 移除零寬度空格。
    text = text.replace("\ufeff", "")  # 移除 BOM 字元。
    text = text.strip()  # 移除前後空格。
    text = " ".join(text.split())  # 將連續空格變成一個空格。
    return text.upper()  # 統一轉為大寫。


@st.cache_data
def read_excel(file_bytes):
    """讀取上傳的 Excel 檔案，並保留原始資料格式。"""
    excel_file = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    return excel_file


def make_excel_download(complete_df, incomplete_df, other_df, diagnostic_df):
    """
    將四個 DataFrame 寫入記憶體中的 Excel 檔案，
    並回傳可直接提供下載的 bytes。
    """
    output = io.BytesIO()  # 建立記憶體檔案。

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # CLEAN 資料寫入「完整資料」。
        complete_df.drop(columns=["_標籤_標準化"], errors="ignore").to_excel(
            writer,
            sheet_name="完整資料",
            index=False
        )

        # DIRTY 資料寫入「不完整資料」。
        incomplete_df.drop(columns=["_標籤_標準化"], errors="ignore").to_excel(
            writer,
            sheet_name="不完整資料",
            index=False
        )

        # 空白或其他標籤寫入「無法識別資料」。
        other_df.drop(columns=["_標籤_標準化"], errors="ignore").to_excel(
            writer,
            sheet_name="無法識別資料",
            index=False
        )

        # 診斷表寫入「分組診斷」。
        diagnostic_df.to_excel(
            writer,
            sheet_name="分組診斷",
            index=False
        )

        workbook = writer.book  # 取得活頁簿物件。
        wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})  # 自動換行格式。

        # 為每個工作表設定凍結窗格、篩選及欄寬。
        for sheet_name, table in {
            "完整資料": complete_df,
            "不完整資料": incomplete_df,
            "無法識別資料": other_df,
            "分組診斷": diagnostic_df
        }.items():
            worksheet = writer.sheets[sheet_name]  # 取得工作表。
            worksheet.freeze_panes(1, 0)  # 固定標題列。
            worksheet.autofilter(0, 0, len(table), len(table.columns) - 1)  # 加入篩選。
            worksheet.set_column(0, len(table.columns) - 1, 18, wrap_format)  # 設定欄寬。

            if sheet_name == "分組診斷":  # 診斷說明較長。
                worksheet.set_column(1, 1, 100, wrap_format)  # 加寬說明欄。

    output.seek(0)  # 回到檔案開頭，準備下載。
    return output.getvalue()  # 回傳 Excel 檔案內容。


# ============================================================
# 上傳 Excel
# ============================================================

uploaded_file = st.file_uploader(
    label="請上傳 Excel 檔案（.xlsx）",
    type=["xlsx"],
    accept_multiple_files=False
)

if uploaded_file is None:
    st.warning("請先上傳一個 .xlsx Excel 檔案。")
    st.stop()  # 未上傳檔案時停止執行後續程式。

st.success(f"已成功上傳：{uploaded_file.name}")  # 顯示上傳檔名。


# ============================================================
# 讀取 Excel
# ============================================================

try:
    excel_file = read_excel(uploaded_file.getvalue())  # 讀取 Excel。
except Exception as error:
    st.error(f"無法讀取 Excel 檔案：{error}")
    st.stop()

st.write("Excel 內的工作表：", excel_file.sheet_names)  # 顯示實際工作表。

if SOURCE_SHEET not in excel_file.sheet_names:
    st.error(
        f"找不到工作表 `{SOURCE_SHEET}`。\n\n"
        f"請確認 Excel 內是否包含此工作表，或修改 `app.py` 頂部的 `SOURCE_SHEET`。"
    )
    st.stop()

try:
    df = pd.read_excel(
        excel_file,
        sheet_name=SOURCE_SHEET,
        dtype=object,  # 保留原始值，避免 pandas 自動改變資料類型。
        keep_default_na=False  # 保留 NULL、N/A 等文字。
    )
except Exception as error:
    st.error(f"讀取工作表 `{SOURCE_SHEET}` 時發生錯誤：{error}")
    st.stop()

if df.empty:
    st.error("來源工作表沒有資料。")
    st.stop()

st.write(f"成功讀取 **{len(df)} 列**、**{len(df.columns)} 欄**。")


# ============================================================
# 尋找 Data_Quality_Label 欄位
# ============================================================

# 建立欄位名稱標準化對照表。
normalized_columns = {
    clean_text(column): column for column in df.columns
}

if clean_text(LABEL_COLUMN) not in normalized_columns:
    st.error(
        f"找不到標籤欄位 `{LABEL_COLUMN}`。\n\n"
        f"Excel 實際欄位名稱：{list(df.columns)}"
    )
    st.stop()

actual_label_column = normalized_columns[clean_text(LABEL_COLUMN)]  # 取得實際欄位名稱。

st.write(f"使用的標籤欄位：`{actual_label_column}`")

# 建立標準化標籤欄位。
df["_標籤_標準化"] = df[actual_label_column].map(clean_text)


# ============================================================
# 顯示原始標籤分布
# ============================================================

label_counts = df["_標籤_標準化"].value_counts(dropna=False)  # 統計標籤數量。

st.subheader("原始 Data_Quality_Label 分布")
st.dataframe(
    label_counts.rename_axis("標準化後標籤").reset_index(name="列數"),
    use_container_width=True
)

clean_count = int(label_counts.get(CLEAN_LABEL, 0))  # CLEAN 數量。
dirty_count = int(label_counts.get(DIRTY_LABEL, 0))  # DIRTY 數量。
other_count = len(df) - clean_count - dirty_count  # 其他或空白標籤數量。

col1, col2, col3 = st.columns(3)  # 建立三欄顯示統計。

col1.metric("CLEAN", clean_count)
col2.metric("DIRTY", dirty_count)
col3.metric("其他／空白", other_count)


# ============================================================
# 嚴格按原始標籤分組
# ============================================================

complete_mask = df["_標籤_標準化"] == CLEAN_LABEL  # CLEAN 的列。
incomplete_mask = df["_標籤_標準化"] == DIRTY_LABEL  # DIRTY 的列。
other_mask = ~(complete_mask | incomplete_mask)  # 其餘列。

complete_df = df.loc[complete_mask].copy()  # 完整資料。
incomplete_df = df.loc[incomplete_mask].copy()  # 不完整資料。
other_df = df.loc[other_mask].copy()  # 無法識別資料。

# 確認三個分組總數等於來源資料總數。
assert len(complete_df) + len(incomplete_df) + len(other_df) == len(df)

st.subheader("分組結果")

result_col1, result_col2, result_col3 = st.columns(3)  # 建立三欄顯示分組結果。

result_col1.metric("完整資料（CLEAN）", len(complete_df))
result_col2.metric("不完整資料（DIRTY）", len(incomplete_df))
result_col3.metric("無法識別資料", len(other_df))

if len(complete_df) == 50 and len(incomplete_df) == 50 and len(other_df) == 0:
    st.success("✅ 已達到預期的 50:50 分組。")
else:
    st.warning(
        "⚠️ 未達到 50:50。請查看下方「分組診斷」，"
        "確認原始標籤是否有空白、大小寫差異、隱藏空格或其他值。"
    )


# ============================================================
# 顯示分組預覽
# ============================================================

st.subheader("資料預覽")

preview_tab1, preview_tab2, preview_tab3 = st.tabs(
    ["完整資料", "不完整資料", "無法識別資料"]
)

with preview_tab1:
    st.dataframe(
        complete_df.drop(columns=["_標籤_標準化"], errors="ignore"),
        use_container_width=True
    )

with preview_tab2:
    st.dataframe(
        incomplete_df.drop(columns=["_標籤_標準化"], errors="ignore"),
        use_container_width=True
    )

with preview_tab3:
    st.dataframe(
        other_df.drop(columns=["_標籤_標準化"], errors="ignore"),
        use_container_width=True
    )


# ============================================================
# 建立分組診斷表
# ============================================================

diagnostic_rows = []  # 儲存診斷資料。

for label, count in label_counts.items():
    diagnostic_rows.append({
        "類型": "標籤統計",
        "內容": f"標準化後標籤 = {label!r}",
        "列數": int(count),
        "是否歸入完整資料": label == CLEAN_LABEL,
        "是否歸入不完整資料": label == DIRTY_LABEL
    })

for index, row in df.head(10).iterrows():
    diagnostic_rows.append({
        "類型": "前 10 列核對",
        "內容": (
            f"Excel 第 {index + 2} 列；"
            f"原始值 = {row[actual_label_column]!r}；"
            f"標準化後 = {row['_標籤_標準化']!r}"
        ),
        "列數": 1,
        "是否歸入完整資料": row["_標籤_標準化"] == CLEAN_LABEL,
        "是否歸入不完整資料": row["_標籤_標準化"] == DIRTY_LABEL
    })

diagnostic_df = pd.DataFrame(diagnostic_rows)  # 轉為診斷表。

st.subheader("分組診斷")
st.dataframe(diagnostic_df, use_container_width=True)


# ============================================================
# 產生並下載 Excel
# ============================================================

st.subheader("下載整理結果")

try:
    excel_bytes = make_excel_download(
        complete_df,
        incomplete_df,
        other_df,
        diagnostic_df
    )
except Exception as error:
    st.error(f"產生 Excel 時發生錯誤：{error}")
    st.stop()

original_stem = Path(uploaded_file.name).stem  # 取得原檔名主檔名。
download_name = f"{original_stem}_最終分組結果.xlsx"  # 產生下載檔名。

st.download_button(
    label="⬇️ 下載整理後的 Excel",
    data=excel_bytes,
    file_name=download_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption(
    "下載檔案包含四個工作表：完整資料、不完整資料、無法識別資料及分組診斷。"
)
