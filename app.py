import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("支援：跨欄置中辨識、自動補零、重複欄位更名")

uploaded_file = st.file_uploader("選擇您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取原始資料 (不設定 header)
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 處理第 2 行跨欄置中的標籤 (向下填充) ---
        # 讓 [KPI, NaN, NaN] 變成 [KPI, KPI, KPI]
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        
        # --- 處理第 4 行標題 ---
        headers = df_raw.iloc[3, :].astype(str)

        # --- 處理數據主體 ---
        data = df_raw.iloc[4:].copy()
        
        # 2. 定義分類索引
        first_col_idx = 0
        business_col_indices = [first_col_idx]
        media_col_indices = [first_col_idx]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i])
            if "KPI" in group_name:
                business_col_indices.append(i)
            elif "Media" in group_name:
                media_col_indices.append(i)

        # 3. 根據索引提取數據 (避免重複名稱導致的遺失)
        df_business = data.iloc[:, business_col_indices]
        df_media = data.iloc[:, media_col_indices]

        # 4. 處理重複欄位名稱的函式
        def handle_duplicates(cols):
            counts = {}
            new_cols = []
            for col in cols:
                if col in counts:
                    counts[col] += 1
                    new_cols.append(f"{col}_{counts[col]}")
                else:
                    counts[col] = 0
                    new_cols.append(col)
            return new_cols

        # 重新指定標題
        df_business.columns = handle_duplicates(headers.iloc[business_col_indices].tolist())
        df_media.columns = handle_duplicates(headers.iloc[media_col_indices].tolist())

        # --- 需求：第一欄日期格式純化 (移除時間) ---
        try:
            date_col = df_business.columns[0]
            df_business[date_col] = pd.to_datetime(df_business[date_col]).dt.date
            df_media[date_col] = pd.to_datetime(df_media[date_col]).dt.date
        except:
            pass

        # --- 需求：所有數值補零 ---
        df_business = df_business.fillna(0)
        df_media = df_media.fillna(0)

        st.success("✅ 檔案處理完成！")

        # --- 下載區 ---
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Business 數據")
            st.download_button(
                label="📥 下載 ROI_Business.xlsx",
                data=to_excel(df_business),
                file_name="ROI_Business.xlsx",
                key="dl_roi_biz"
            )
            st.dataframe(df_business.head(10))

        with col2:
            st.subheader("Media 數據")
            st.download_button(
                label="📥 下載 ROI_Media.xlsx",
                data=to_excel(df_media),
                file_name="ROI_Media.xlsx",
                key="dl_roi_media"
            )
            st.dataframe(df_media.head(10))

    except Exception as e:
        st.error(f"轉換過程中發生錯誤：{e}")
