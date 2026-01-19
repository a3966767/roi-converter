import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v5.0", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("校準：1.排除 Non-media 2.字串精確替換 (imp->Impressions等) 3.特殊媒體排序 4.徹底修正 Arg 錯誤")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy().reset_index(drop=True)

        # --- Business 處理 ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        for col in df_biz.columns[1:]:
            df_biz[col] = pd.to_numeric(df_biz[col], errors='coerce').fillna(0)

        # --- Media 處理 ---
        # 排除 Non-media
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and "NON MEDIA" not in str(g).upper() and i > 0]
        
        special_keywords = ["OWNED MEDIA", "SHARED MEDIA", "EARNED MEDIA"]
        
        # 建立欄位規則映射
        col_mapping = {}
        normal_cats, special_cats = [], []

        for i in media_idx:
            if i == 0: continue
            g_name = str(groups[i]).upper()
            h_name = raw_headers[i]
            parts = h_name.split("_")
            
            is_special = any(k in g_name for k in special_keywords)
            
            if is_special:
                # 規則 2: 取第一個底線前
                cat = parts[0].upper()
                if cat not in special_cats: special_cats.append(cat)
                col_mapping[i] = {"cat": cat, "type": "special"}
            else:
                # 規則 1: 取最後一個底線前
                cat = "_".join(parts[:-1]).upper() if len(parts) > 1 else h_name.upper()
                if cat not in normal_cats: normal_cats.append(cat)
                col_mapping[i] = {"cat": cat, "type": "normal"}

        # 排序：一般在前，特殊在後
        final_cat_order = normal_cats + [c for c in special_cats if c not in normal_cats]
        
        # 定義精確替換辭典 (規則 3)
        rename_dict = {
            "imp": "Impressions",
            "view": "Views",
            "click": "Clicks",
            "spend": "Spend",
            "spent": "Spend",
            "grp": "GRP"
        }

        all_chunks = []
        for cat in final_cat_order:
            target_indices = [i for i, info in col_mapping.items() if info["cat"] == cat]
            if not target_indices: continue
            
            # 建立 DataFrame 並確保每一欄都是單獨選取的 Series (避免 Arg Error)
            temp_df = pd.DataFrame()
            # 處理日期
            date_series = pd.to_datetime(data.iloc[:, 0], errors='coerce').dt.strftime('%Y-%m-%d')
            temp_df["Date"] = date_series
            temp_df["Media"] = cat
            temp_df["Product"] = "illuma"
            
            for i in target_indices:
                h_name = raw_headers[i]
                parts = h_name.split("_")
                raw_sub_h = parts[-1].lower() # 先轉小寫以便比對辭典
                
                # 精確替換關鍵字，若不在辭典內則字首大寫
                clean_h = rename_dict.get(raw_sub_h, raw_sub_h.capitalize())
                
                # 數值轉型
                val_series = pd.to_numeric(data.iloc[:, i], errors='coerce').fillna(0)
                temp_df[clean_h] = val_series
                
            all_chunks.append(temp_df)

        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        st.success("✅ 轉換成功！已套用精確文字替換與排序。")

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
        st.error(f"❌ 處理失敗：{e}")
