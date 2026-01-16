import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據自動轉換工具 v2.2", layout="centered")

st.title("📊 ROI 數據自動分類轉換器 v2.2")
st.write("修正：支援跨欄置中標籤自動延伸 & 數值補零")

uploaded_file = st.file_uploader("選擇您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取原始資料 (不設 header)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 處理第 2 行跨欄置中的問題 ---
        # 提取第 2 行 (索引為 1)，並使用 ffill() 讓標籤向後延伸
        # 例如: [KPI, NaN, NaN, Media] -> [KPI, KPI, KPI, Media]
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        
        # 提取第 4 行標題與數據
        headers = df_raw.iloc[3, :]
        data = df_raw.iloc[4:].copy()
        data.columns = headers

        # --- 需求 1: Date/Week 格式純化 ---
        first_col_name = headers.iloc[0]
        try:
            data[first_col_name] = pd.to_datetime(data[first_col_name]).dt.strftime('%Y-%m-%d')
        except:
            pass

        # --- 需求 2 & 3: 分類邏輯 ---
        business_cols = [first_col_name]
        media_cols = [first_col_name]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i])
            col_name = headers.iloc[i]
            
            # 如果這一欄屬於 KPI 群組 (包含跨欄置中延伸過來的)
            if "KPI" in group_name:
                business_cols.append(col_name)
            # 如果這一欄屬於 Media 群組
            elif "Media" in group_name:
                media_cols.append(col_name)

        # 需求 4: 建立資料表並強制補零
        df_business = data[business_cols].fillna(0)
        df_media = data[media_cols].fillna(0)

        st.success("檔案解析成功！跨欄標籤已自動辨識。")

        # --- 下載區 ---
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Business 檔案")
            st.download_button("📥 下載 Business", to_excel(df_business), "ROI_Business.xlsx", key="dl_biz")
            st.dataframe(df_business.head())

        with col2:
            st.subheader("Media 檔案")
            st.download_button("📥 下載 Media", to_excel(df_media), "ROI_Media.xlsx", key="dl_media")
            st.dataframe(df_media.head())

    except Exception as e:
        st.error(f"處理失敗：{e}")
