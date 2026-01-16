import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據自動轉換工具 v2.1", layout="centered")

st.title("📊 ROI 數據自動分類轉換器 v2.1")
st.write("已修正預覽語法錯誤，支援補零與日期純化")

uploaded_file = st.file_uploader("選擇您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取原始資料
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 提取資訊
        groups = df_raw.iloc[1, :]   # 第二行：分類群組
        headers = df_raw.iloc[3, :]  # 第四行：變數名稱
        data = df_raw.iloc[4:].copy() # 第五行以後：數值資料
        data.columns = headers       # 設定欄位名

        # --- 需求 1: Week 部分不需要時間 ---
        first_col_name = headers.iloc[0]
        # 轉換日期並確保不帶時間分秒
        try:
            data[first_col_name] = pd.to_datetime(data[first_col_name]).dt.strftime('%Y-%m-%d')
        except:
            pass # 如果第一欄不是日期格式則跳過，避免報錯

        # --- 需求 2 & 3: 分類與補零邏輯 ---
        business_cols = [first_col_name]
        media_cols = [first_col_name]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i])
            col_name = headers.iloc[i]
            
            # 需求 2: 含有 KPI 歸類為 Business
            if "KPI" in group_name:
                business_cols.append(col_name)
            # 需求 3: 含有 Media 歸類為 Media
            elif "Media" in group_name:
                media_cols.append(col_name)

        # 建立資料表並補零 (需求 2 & 4)
        df_business = data[business_cols].fillna(0)
        df_media = data[media_cols].fillna(0)

        st.success("檔案處理成功！")

        # --- 下載區 ---
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Business 檔案")
            st.download_button(
                label="📥 下載 ROI_Business.xlsx",
                data=to_excel(df_business),
                file_name="ROI_Business.xlsx",
                key="dl_biz"
            )
            st.write("數據預覽：")
            st.dataframe(df_business.head())

        with col2:
            st.subheader("Media 檔案")
            st.download_button(
                label="📥 下載 ROI_Media.xlsx",
                data=to_excel(df_media),
                file_name="ROI_Media.xlsx",
                key="dl_media"
            )
            st.write("數據預覽：")
            st.dataframe(df_media.head())

    except Exception as e:
        st.error(f"處理檔案時出錯：{e}")
