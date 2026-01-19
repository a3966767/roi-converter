import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.5", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("邏輯更新：區分一般媒體與 Owned/Shared/Earned 處理規則，標題首字母大寫並重新排序。")

uploaded_file = st.file_uploader("選擇原始檔案", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- Business 處理 ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- Media 處理 ---
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and i > 0]
        
        # 定義特殊關鍵字
        special_keywords = ["OWNED MEDIA", "SHARED MEDIA", "EARNED MEDIA"]
        
        # 建立分類桶
        normal_cats = []   # 一般媒體
        special_cats = []  # Owned/Shared/Earned
        
        col_to_info = {}
        for i in media_idx:
            if i == 0: continue
            g_name = str(groups[i]).upper()
            h_name = raw_headers[i]
            
            if any(k in g_name for k in special_keywords):
                # 規則 2: 特殊分類，取第一個底線以前
                cat = h_name.split("_")[0].upper()
                clean_h = h_name # 特殊分類標題保留或後續處理
                if cat not in special_cats: special_cats.append(cat)
                col_to_info[i] = {"cat": cat, "is_special": True}
            else:
                # 規則 1: 一般媒體，取最後底線之前作為 Media
                parts = h_name.split("_")
                cat = "_".join(parts[:-1]).upper() if len(parts) > 1 else h_name.upper()
                if cat not in normal_cats: normal_cats.append(cat)
                col_to_info[i] = {"cat": cat, "is_special": False}

        # 合併排序：一般在前，特殊在後
        final_cat_order = normal_cats + special_cats
        
        all_chunks = []
        for cat in final_cat_order:
            sub_indices = [0]
            sub_headers = ["Date"]
            
            is_cat_special = cat in special_cats
            
            for i in media_idx:
                if i == 0: continue
                if col_to_info[i]["cat"] == cat:
                    sub_indices.append(i)
                    h_name = raw_headers[i]
                    
                    # 規則 3 & 4: 處理標題文字
                    parts = h_name.split("_")
                    raw_sub_h = parts[-1] if len(parts) > 1 else h_name
                    # 首字母大寫轉換
                    clean_h = raw_sub_h.capitalize()
                    sub_headers.append(clean_h)
            
            df_temp = data.iloc[:, sub_indices].copy()
            df_temp.columns = sub_headers
            
            # 插入 B, C 欄
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 類型轉換
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
            for col in df_temp.columns[3:]:
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
                
            all_chunks.append(df_temp)

        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        st.success("✅ 規則轉換完成！")
        
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 處理失敗：{e}")
