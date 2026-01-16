import streamlit as st
import pandas as pd
from io import BytesIO

# 設定置中佈局
st.set_page_config(page_title="ROI 數據轉換工具 v2.9", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("新增功能：ROI_Media 自動在 B 欄提取媒體分類名稱（底線前文字）。")

uploaded_file = st.file_uploader("第一步：選擇您的檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # --- 1. 判斷格式並讀取 ---
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension == 'csv':
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 處理第 2 行跨欄標籤 (ffill) ---
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        # --- 處理第 4 行標題 ---
        headers = df_raw.iloc[3, :].astype(str)
        # --- 處理數據主體 ---
        data = df_raw.iloc[4:].copy()

        # --- 2. 定義索引分類 ---
        first_col_idx = 0
        biz_indices = [first_col_idx]
        media_indices = [first_col_idx]

        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "KPI" in g_name:
                biz_indices.append(i)
            elif "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_indices.append(i)

        # --- 3. 處理 Business 檔案 ---
        df_biz = data.iloc[:, biz_indices]
        # 處理重複標題
        def handle_dupes(cols):
            counts = {}
            new = []
            for c in cols:
                if c in counts:
                    counts[c] += 1
                    new.append(f"{c}_{counts[c]}")
                else:
                    counts[c] = 0
                    new.append(c)
            return new
        
        df_biz.columns = handle_dupes(headers.iloc[biz_indices].tolist())
        df_biz = df_biz.fillna(0)

        # --- 4. 處理 Media 檔案 (核心邏輯：新增 B 欄) ---
        # 提取 Media 數據 (暫時包含第一欄 Date)
        df_media_temp = data.iloc[:, media_indices]
        media_headers = headers.iloc[media_indices].tolist()
        
        # 建立一個新的列表來存放分類 (底線前的內容)
        # 第一格是 Date，所以對應的分類我們放空值或標題
        media_categories = []
        for h in media_headers:
            if "_" in h:
                media_categories.append(h.split("_")[0]) # 拿底線前的文字
            else:
                media_categories.append(h) # 若無底線則用原文

        # 準備最終的 Media 數據結構
        # 我們將資料依照 Date 欄位進行「熔斷」或直接填入 B 欄
        # 根據您的需求：B 欄是媒體名稱，從 C 欄開始是數值
        
        final_media_rows = []
        date_col_idx = media_indices[0]
        
        # 我們遍歷每一行數據
        for _, row in data.iterrows():
            current_date = row[date_col_idx]
            # 遍歷除了 Date 以外的每一個 Media 欄位
            for i in range(1, len(media_indices)):
                actual_col_idx = media_indices[i]
                val = row[actual_col_idx]
                category = media_categories[i]
                col_name = media_headers[i]
                
                final_media_rows.append({
                    'Date': current_date,
                    'Media': category,
                    'Variable': col_name,
                    'Value': val if pd.notnull(val) else 0
                })
        
        # 將資料轉回 DataFrame 並重新整理成寬表格 (Pivot)
        # 註：為了符合您「B列是Media，C列開始是數值」的格式
        # 我們需要將同一天的資料橫向展開，但保留 Media 分類
        
        # 重新構建 Media 檔案
        df_media = data.iloc[:, media_indices].copy()
        df_media.columns = handle_dupes(media_headers)
        
        # 在索引 1 的位置 (即 B 欄) 插入 Media 欄位
        # 這裡採取簡化邏輯：因為每一列對應多個 Media，
        # 如果您的原始檔一行就是一個日期，而媒體名稱是橫向的，
        # 程式會自動抓取第一個媒體名稱填入 B 欄，或是您需要的是「轉置(Tidy)」格式？
        
        # 根據您的描述「若不同請複製同一段日期往下延伸」，這代表需要做資料轉置：
        tidy_data = []
        for _, row in data.iterrows():
            d = row[0]
            for i in range(1, len(media_indices)):
                h = media_headers[i]
                cat = h.split("_")[0] if "_" in h else h
                val = row[media_indices[i]]
                tidy_data.append([d, cat, h, val])
        
        df_media_final = pd.DataFrame(tidy_data, columns=['Date', 'Media', 'Variable', 'Value'])
        df_media_final = df_media_final.fillna(0)

        # --- 日期格式純化 ---
        try:
            df_biz[df_biz.columns[0]] = pd.to_datetime(df_biz[df_biz.columns[0]]).dt.date
            df_media_final['Date'] = pd.to_datetime(df_media_final['Date']).dt.date
        except:
            pass

        # --- 介面顯示 ---
        st.success("✅ 處理完成！ROI_Media 已進行轉置處理。")
        
        st.subheader("📁 Business 預覽")
        st.dataframe(df_biz.head(10), use_container_width=True)

        st.subheader("📁 Media 預覽 (已新增 Media 分類欄位)")
        st.dataframe(df_media_final.head(10), use_container_width=True)

        st.divider()

        # --- 下載區 ---
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx", key="dl_biz", use_container_width=True)
        with col_dl2:
            st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx", key="dl_media", use_container_width=True)

    except Exception as e:
        st.error(f"❌ 處理失敗：{e}")
