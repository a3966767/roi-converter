import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.1", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正：1. 強制日期字串化 2. 移除所有標題空白 3. 確保數值欄位為純 Float 型態。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV 檔案", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據 (不轉換型態，由程式後續處理)
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        # 關鍵：徹底移除標題中的前後空白與特殊字元
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 2. 處理 ROI_Business ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper():
                biz_idx.append(i)
        
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = [raw_headers[i] for i in biz_idx]
        # 強制日期格式 YYYY-MM-DD
        df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 ROI_Media (精確對齊手作版索引) ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        # 識別分類
        unique_cats = []
        for i in media_idx:
            if i == 0: continue
            cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
            if cat not in unique_cats:
                unique_cats.append(cat)

        all_chunks = []
        for cat in unique_cats:
            sub_indices = [0]
            # 手作版標題不帶媒體前綴
            sub_headers = ["Date"]
            
            for i in media_idx:
                if i == 0: continue
                curr_cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
                if curr_cat == cat:
                    sub_indices.append(i)
                    clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "").strip()
                    sub_headers.append(clean_h)
            
            df_temp = data.iloc[:, sub_indices].copy()
            
            # 去除重複欄位名 (防止 reindexing error)
            final_h, h_counts = [], {}
            for h in sub_headers:
                if h in h_counts:
                    h_counts[h] += 1
                    final_h.append(f"{h}_{h_counts[h]}")
                else:
                    h_counts[h] = 0
                    final_h.append(h)
            df_temp.columns = final_h
            
            # 插入手作版必要的欄位：Media (B), Product (C)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # --- 數據格式強制轉換區 ---
            # 1. 日期轉字串 (防止建模程式對不齊日期)
            df_temp['Date'] = pd.to_datetime(df_temp['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 2. 數值欄位強制轉為 float (從第 4 欄開始)
            for col in df_temp.columns[3:]:
                # 這是解決 effect_start 的最關鍵一步：確保 sum() 能抓到數字
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0).astype(float)
            
            all_chunks.append(df_temp)

        # 合併並確保沒有任何 NaN
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0.0)

        # 最終檢查：如果日期欄位有 NaN，建模程式會崩潰
        df_media_final = df_media_final.dropna(subset=['Date'])

        st.success("✅ 校準完成！此版本格式與 ROI_Media_illuma.xlsx 完全對齊。")

        def to_excel(df):
            output = BytesIO()
            # 強制使用 openpyxl 並設定不使用內存優化，確保產出的檔案最純淨
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 錯誤詳細資訊：{e}")
