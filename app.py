import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.0", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正說明：確保數據轉換函數僅作用於單一欄位，解決 arg must be a list 報錯。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV 檔案", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 處理第 2 行跨欄標籤 (ffill) 與 第 4 行標題
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 2. 處理 ROI_Business ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper():
                biz_idx.append(i)
        
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = [raw_headers[i] for i in biz_idx]
        
        # 安全日期轉換 (僅針對第一欄)
        df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 ROI_Media (核心邏輯優化) ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        # 找出所有媒體分類
        unique_cats = []
        for i in media_idx:
            if i == 0: continue
            cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
            if cat not in unique_cats:
                unique_cats.append(cat)

        all_chunks = []
        for cat in unique_cats:
            # 建立子區塊索引與標題
            sub_indices = [0]
            sub_headers = ["Date"]
            
            for i in media_idx:
                if i == 0: continue
                # 判斷是否屬於當前媒體分類
                current_col_cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
                if current_col_cat == cat:
                    sub_indices.append(i)
                    # 清理標題，移除媒體前綴
                    clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "")
                    sub_headers.append(clean_h)
            
            # 提取數據並賦予唯一標題
            df_temp = data.iloc[:, sub_indices].copy()
            
            # 處理重複標題 (去重)
            final_headers = []
            counts = {}
            for h in sub_headers:
                if h in counts:
                    counts[h] += 1
                    final_headers.append(f"{h}_{counts[h]}")
                else:
                    counts[h] = 0
                    final_headers.append(h)
            df_temp.columns = final_headers
            
            # 插入 B 欄 Media 與 C 欄 Product (模擬手作版)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 【關鍵修正點】強制類型轉換 (逐欄處理，避開 arg must be a list 報錯)
            # 1. 日期欄位
            df_temp['Date'] = pd.to_datetime(df_temp['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            # 2. 數值欄位 (從第 4 欄起)
            for col in df_temp.columns[3:]:
                # 強制轉換為 Series 後再轉為 numeric
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            all_chunks.append(df_temp)

        # 合併所有區塊
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0)

        st.success("✅ 轉換成功！已解決數據類型轉換報錯。")

        # 預覽與下載
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.subheader("📁 Business 預覽")
        st.dataframe(df_biz.head(5))
        st.subheader("📁 Media 預覽")
        st.dataframe(df_media_final.head(10))

        st.divider()
        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 錯誤詳細資訊：{e}")
