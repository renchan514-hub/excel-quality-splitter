# Excel 資料品質分組工具

這是一個 Streamlit 網頁應用程式，用於上傳 Excel 檔案，並根據 `Data_Quality_Label` 欄位分組資料。

## 功能

- 上傳 `.xlsx` Excel 檔案
- 讀取 `Dataset_Combined` 工作表
- 將 `CLEAN` 資料放入「完整資料」
- 將 `DIRTY` 資料放入「不完整資料」
- 將空白、大小寫異常或其他標籤放入「無法識別資料」
- 顯示標籤分布、分組統計及診斷資訊
- 下載包含四個工作表的整理結果 Excel

## 使用方法

1. 開啟 Streamlit App。
2. 上傳 `.xlsx` Excel 檔案。
3. 檢查 `Data_Quality_Label` 分布及分組結果。
4. 按「下載整理後的 Excel」取得結果。

## 部署

本 App 可部署至 Streamlit Community Cloud。
