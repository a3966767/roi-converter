import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.3", layout="centered")

st.title("📊 ROI 數據自動分類轉換器 (手作模擬版)")
st.write("此版本產出的格式與您手作的 ROI_Media_illuma.xlsx 完全一致。")

uploaded_file = st.file_uploader("上傳原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        headers = df_raw.iloc[3, :].astype(str).tolist()
        data = df_raw.iloc[4:].copy()

        # --- Business 處理 (不補 0，保持空白) ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper(): biz_idx.append(i)
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = [headers[i] for i in biz_idx]

        # --- Media 處理 (核心優化) ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        media_categories = []
        for i in media_idx:
            h = headers[i]
            # 模擬手作版：底線前文字並轉大寫
            media_categories.append(h.split("_")[0].upper() if "_" in h else h.upper())

        unique_cats = [c for c in list(dict.fromkeys(media_categories)) if c != "DATE"]
        all_chunks = []
        
        for cat in unique_cats:
            cat_sub_idx = [media_idx[0]]
            cat_sub_head = ["Date"] # 固定第一欄叫 Date
            
            for i in range(1, len(media_idx)):
                if media_categories[i] == cat:
                    cat_sub_idx.append(media_idx[i])
                    # 模擬手作版：移除標題中的媒體前綴，讓標題變乾淨
                    clean_head = headers[media_idx[i]].replace(f"{cat.lower()}_", "").replace(f"{cat}_", "")
                    cat_sub_head.append(clean_head)
            
            df_temp = data.iloc[:, cat_sub_idx].copy()
            df_temp.columns = cat_sub_head
            
            # 關鍵：不使用 fillna(0)，保持為空白，與手作版一致
            df_temp.insert(1, 'Media', cat)
            all_chunks.append(df_temp)

        # 合併數據
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        # 日期處理
        try:
            df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0]).dt.date
            df_media_final.iloc[:, 0] = pd.to_datetime(df_media_final.iloc[:, 0]).dt.date
        except: pass

        st.success("✅ 格式已校準為手作版樣式！")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        c1, c2 = st.columns(2)
        with c1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"錯誤：{e}")
