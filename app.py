import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.8", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("校準重點：1.強迫日期轉為字串 2.加入空白欄位對齊索引 3.強制數值轉換。")

uploaded_file = st.file_uploader("選擇原始檔案 (xlsx/csv)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        # 讀取時確保日期欄位不會被自動轉換成奇怪的格式
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 1. Business 處理 ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper(): biz_idx.append(i)
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = [raw_headers[i] for i in biz_idx]
        df_biz = df_biz.fillna(0)
        # 關鍵：強迫日期變為 YYYY-MM-DD 字串
        df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0]).dt.strftime('%Y-%m-%d')

        # --- 2. Media 處理 ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        col_to_cat = {}
        for i in media_idx:
            if i == 0: col_to_cat[i] = "DATE"
            else:
                h = raw_headers[i]
                col_to_cat[i] = h.split("_")[0].upper() if "_" in h else h.upper()

        unique_cats = [c for c in list(dict.fromkeys(col_to_cat.values())) if c != "DATE"]
        all_chunks = []
        
        for cat in unique_cats:
            cat_sub_idx = [0]
            cat_sub_head = ["Date"]
            for i in media_idx:
                if i == 0: continue
                if col_to_cat[i] == cat:
                    cat_sub_idx.append(i)
                    clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "")
                    cat_sub_headers = clean_h
                    cat_sub_head.append(cat_sub_headers)
            
            df_temp = data.iloc[:, cat_sub_idx].copy()
            df_temp.columns = cat_sub_head
            
            # 插入兩列：Media 與 空白 (對應手作版的 Media & Product)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Placeholder', '') # 對齊手作版的 C 欄，確保數據從 D 開始
            
            # 強制數值轉型，確保 sum() 運算正常
            for col in df_temp.columns[3:]:
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            # 強迫日期變為 YYYY-MM-DD 字串
            df_temp.iloc[:, 0] = pd.to_datetime(df_temp.iloc[:, 0]).dt.strftime('%Y-%m-%d')
            all_chunks.append(df_temp)

        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        st.success("✅ 校準完成！")
        
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
