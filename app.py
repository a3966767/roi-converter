import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.2", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正說明：自動偵測日期欄位，防止 KeyError: 'Date' 錯誤。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV 檔案", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 2. 處理 ROI_Business ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper():
                biz_idx.append(i)
        
        df_biz = data.iloc[:, biz_idx].copy()
        # 強制將第一欄命名為 Date，防止 KeyError
        biz_cols = [raw_headers[i] for i in biz_idx]
        biz_cols[0] = "Date"
        df_biz.columns = biz_cols
        
        # 安全日期轉換
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 ROI_Media ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        unique_cats = []
        for i in media_idx:
            if i == 0: continue
            cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
            if cat not in unique_cats:
                unique_cats.append(cat)

        all_chunks = []
        for cat in unique_cats:
            sub_indices = [0]
            sub_headers = ["Date"] # 強制設定第一個標題為 Date
            
            for i in media_idx:
                if i == 0: continue
                current_col_cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
                if current_col_cat == cat:
                    sub_indices.append(i)
                    clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "").strip()
                    sub_headers.append(clean_h)
            
            df_temp = data.iloc[:, sub_indices].copy()
            
            # 處理重複標題
            final_h, counts = [], {}
            for h in sub_headers:
                if h in counts:
                    counts[h] += 1
                    final_h.append(f"{h}_{counts[h]}")
                else:
                    counts[h] = 0
                    final_h.append(h)
            df_temp.columns = final_h
            
            # 插入 B, C 欄 (模擬手作版)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 【防呆修正】確保對存在的 Date 欄位進行轉換
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 數值欄位轉換
            for col in df_temp.columns[3:]:
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            all_chunks.append(df_temp)

        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0)

        st.success("✅ 轉換成功！已修正 Date 欄位識別問題。")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 錯誤詳細資訊：{e}")
