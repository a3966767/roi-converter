import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換穩定版", layout="centered")

st.title("📊 ROI 數據自動分類轉換器 v2.3")
st.write("修正：解決數值複製不完整與欄位名稱重複問題")

uploaded_file = st.file_uploader("選擇您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取原始資料 (不設定 header，確保完整讀入)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 處理第 2 行跨欄置中的標籤 (填充) ---
        # fillna(method='ffill') 讓 [KPI, NaN, NaN] 變成 [KPI, KPI, KPI]
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        
        # --- 處理第 4 行標題 ---
        headers = df_raw.iloc[3, :]
        
        # --- 處理數據主體 (從第 5 行開始到最後一行) ---
        data = df_raw.iloc[4:].copy()
        
        # 重要修正：先用「數字索引」來分類，避免重複標題造成的數據遺失
        first_col_idx = 0
        business_col_indices = [first_col_idx]
        media_col_indices = [first_col_idx]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i])
            # 判斷邏輯
            if "KPI" in group_name:
                business_col_indices.append(i)
            elif "Media" in group_name:
                media_col_indices.append(i)

        # 2. 根據「位置」取出數據
        df_business = data.iloc[:, business_col_indices]
        df_media = data.iloc[:, media_col_indices]

        # 3. 重新賦予正確的標題
        df_business.columns = headers.iloc[business_col_indices]
        df_media.columns = headers.iloc[media_col_indices]

        # --- 需求 1: 第一欄 Date 格式純化 ---
        try:
            # 轉換第一欄 (Week/Date)
            date_col = df_business.columns[0]
            df_business[date_col] = pd.to_datetime(df_business[date_col]).dt.date
            df_media[date_col] = pd.to_datetime(df_media[date_col]).dt.date
        except:
            pass

        # 需求 2: 全域補零 (包含空白處)
        df_business = df_business.fillna(0)
        df_media = df_media.fillna(0)

        st.success(f"解析成功！Business 欄位數：{len(df_business.columns)}，Media 欄位數：{len(df_media.columns)}")

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
        st.error(f"處理失敗，錯誤原因：{e}")
