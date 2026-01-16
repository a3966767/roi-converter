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
            return new
