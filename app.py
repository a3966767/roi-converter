import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.8", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("規則更新：1. 修正 arg 報錯 2. 一般媒體取末尾 3. 特殊媒體取首段 4. 排序優化")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 提取結構
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy().reset_index(drop=True)

        # --- 處理 Business ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        
        # 轉換 Business 數值
        for col in df_biz.columns[1:]:
            df_biz[col] = pd.to_numeric(df_biz[col], errors='coerce').fillna(0)

        # --- 處理 Media ---
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and i > 0]
        special_keywords = ["OWNED MEDIA", "SHARED MEDIA", "EARNED MEDIA"]
        
        normal_data_list = []
        special_data_list = []
        
        # 取得所有分類名稱
        all_cats_info = {}
        for i in media_idx:
            if i == 0: continue
            g_name = str(groups[i]).upper()
            h_name = raw_headers[i]
            parts = h_name.split("_")
            
            is_special = any(k in g_name for k in special_keywords)
            
            if is_special:
                cat = parts[0].upper()
                info_type = "special"
            else:
                cat = "_".join(parts[:-1]).upper() if len(parts) > 1 else h_name.upper()
                info_type = "normal"
            
            all_cats_info[i] = {"cat": cat, "type": info_type}

        # 取得不重複的分類並排序 (一般在前，特殊在後)
        unique_normal = []
        for i, info in all_cats_info.items():
            if info["type"] == "normal" and info["cat"] not in unique_normal:
                unique_normal.append(info["cat"])
        
        unique_special = []
        for i, info in all_cats_info.items():
            if info["type"] == "special" and info["cat"] not in unique_special:
                unique_special.append(info["cat"])
        
        sorted_cats = unique_normal + unique_special

        all_media_chunks = []
        for cat in sorted_cats:
            # 找出屬於此 cat 的原始索引
            cat_indices = [i for i, info in all_cats_info.items() if info["cat"] == cat]
            
            # 建立該分類的數據
            temp_df = pd.DataFrame()
            temp_df["Date"] = pd.to_datetime(data.iloc[:, 0], errors='coerce').dt.strftime('%Y-%m-%d')
            temp_df["Media"] = cat
            temp_df["Product"] = "illuma"
            
            for i in cat_indices:
                h_name = raw_headers[i]
                parts = h_name.split("_")
                raw_sub_h = parts[-1] if len(parts) > 1 else h_name
                
                # 規範化標題
                clean_h = raw_sub_h.capitalize()
                if clean_h == "Spent": clean_h = "Spend"
                
                # 轉型並加入
                temp_df[clean_h] = pd.to_numeric(data.iloc[:, i], errors='coerce').fillna(0)
            
            all_media_chunks.append(temp_df)

        df_media_final = pd.concat(all_media_chunks, axis=0, ignore_index=True)

        st.success("✅ 轉換成功！已解決 arg 報錯並完成排序。")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col1, col2 = st.columns(2)
        with col1: st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with col2: st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 處理失敗：{e}")
