import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據自動轉換工具", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.write("請上傳原始 Excel 檔，程式將自動根據第二行分類並提供下載。")

# --- 上傳檔案 ---
uploaded_file = st.file_uploader("選擇您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取原始資料 (不設 header，因為我們要手動解析前幾行)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 提取資訊：第 2 行是群組，第 4 行是欄位名
        groups = df_raw.iloc[1, :]
        headers = df_raw.iloc[3, :]
        data = df_raw.iloc[4:].copy()
        data.columns = headers # 重新定義欄位名稱

        # 2. 分類邏輯
        # 假設第一欄通常是 Date 或 ID，兩邊都保留
        business_cols = [headers.iloc[0]]
        media_cols = [headers.iloc[0]]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i])
            col_name = headers.iloc[i]
            
            # 判斷邏輯：含 KPI 字眼的歸類為 Business，其餘歸類為 Media
            if "KPI" in group_name:
                business_cols.append(col_name)
            else:
                media_cols.append(col_name)

        df_business = data[business_cols]
        df_media = data[media_cols]

        st.success("檔案解析成功！")

        # --- 顯示預覽 ---
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Business 預覽")
            st.dataframe(df_business.head())
        with col2:
            st.subheader("Media 預覽")
            st.dataframe(df_media.head())

        # --- 下載按鈕函數 ---
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        st.subheader("📥 下載轉化後的檔案")
        
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label="下載 ROI_Business.xlsx",
                data=to_excel(df_business),
                file_name="ROI_Business.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with dl_col2:
            st.download_button(
                label="下載 ROI_Media.xlsx",
                data=to_excel(df_media),
                file_name="ROI_Media.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"處理檔案時出錯：{e}")