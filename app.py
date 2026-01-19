import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.6", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("校準重點：強制數值類型轉換 (Float)，確保建模程式 sum() 運算正常。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).tolist()
        data = df_raw.iloc[4:].copy()

        def make_unique(cols):
            counts = {}
            new_cols = []
            for col in cols:
                c = str(col).strip()
                if c in counts:
                    counts[c] += 1
                    new_cols.append(f"{c}_{counts[c]}")
                else:
                    counts[c] = 0
                    new_cols.append(c)
            return new_cols

        # --- 2. 處理 ROI_Business ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper(): biz_idx.append(i)
        df_biz = data.iloc[:, biz_idx].copy()
        biz_header_list = [raw_headers[i] for i in biz_idx]
        biz_header_list[0] = "Date"
        df_biz.columns = make_unique(biz_header_list)
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 ROI_Media (B欄 Media, C欄起為數據) ---
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
        
        all_media_chunks = []
        for cat in unique_cats:
            cat_sub_idx = [0]
            cat_sub_headers = ["Date"]
            
            for i in media_idx:
                if i == 0: continue
                if col_to_cat[i] == cat:
                    cat_sub_idx.append(i)
                    orig_h = raw_headers[i]
                    # 移除前綴，保持標題乾淨 (如 Impressions, Spent)
                    clean_h = orig_h.replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "")
                    cat_sub_headers.append(clean_h)
            
            df_temp = data.iloc[:, cat_sub_idx].copy()
            df_temp.columns = make_unique(cat_sub_headers)
            
            # 插入 Media 欄位在 B 欄
            df_temp.insert(1, 'Media', cat)
            
            # --- 關鍵修正：強制轉換 C 欄以後為 Float (浮點數) ---
            # 這樣可以確保 sum() 運算時 0.0 + 0.0 是有效的數字運算
            for col in df_temp.columns[2:]: 
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
                
            all_media_chunks.append(df_temp)

        # 合併數據
        df_media_final = pd.concat(all_media_chunks, axis=0, ignore_index=True)

        # 日期格式化
        try:
            df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0]).dt.date
            df_media_final.iloc[:, 0] = pd.to_datetime(df_media_final.iloc[:, 0]).dt.date
        except: pass

        st.success("✅ 轉換完成！已強制數值類型校準。")
        
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx", use_container_width=True)
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx", use_container_width=True)

    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")
