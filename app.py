import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 ROI 數據轉換器 (終極相容版 v4.3)")

uploaded_file = st.file_uploader("上傳原始檔", type=["xlsx", "csv"])

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 1. Business 處理 ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        # 強制日期格式化為 YYYY-MM-DD 字串
        df_biz["Date"] = pd.to_datetime(df_biz["Date"]).dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 2. Media 處理 ---
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and "NON MEDIA" not in str(g).upper() and i > 0]
        
        unique_cats = []
        for i in media_idx[1:]:
            cat = raw_headers[i].split("_")[0].upper()
            if cat not in unique_cats: unique_cats.append(cat)

        all_chunks = []
        for cat in unique_cats:
            sub_indices = [0]
            sub_headers = ["Date"]
            for i in media_idx[1:]:
                curr_cat = raw_headers[i].split("_")[0].upper()
                if curr_cat == cat:
                    sub_indices.append(i)
                    sub_headers.append(raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "").strip())
            
            df_temp = data.iloc[:, sub_indices].copy()
            df_temp.columns = sub_headers
            # 插入 B, C 欄 (完全模擬手作版)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 關鍵：日期強制與 Business 一致
            df_temp["Date"] = pd.to_datetime(df_temp["Date"]).dt.strftime('%Y-%m-%d')
            
            # 關鍵：數值強制轉換，確保 sum() > 0 會生效
            for col in df_temp.columns[3:]:
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            all_chunks.append(df_temp)

        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.success("✅ 格式校準完成！請下載後直接放入建模程式。")
        st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"錯誤：{e}")
