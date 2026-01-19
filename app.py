import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.9", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正：解決 arg must be a list 錯誤，確保數據類型轉換完全符合 Pandas 規範。")

uploaded_file = st.file_uploader("選擇原始檔案 (xlsx/csv)", type=["xlsx", "csv"])

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

        # --- 2. Business 處理 ---
        biz_idx = [0]
        for i in range(1, len(groups)):
            if "KPI" in str(groups[i]).upper():
                biz_idx.append(i)
        
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = [raw_headers[i] for i in biz_idx]
        
        # 安全轉換日期
        df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 3. Media 處理 ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        # 建立分類對照
        unique_cats = []
        col_to_cat = {}
        for i in media_idx:
            if i == 0: continue
            cat = raw_headers[i].split("_")[0].upper() if "_" in raw_headers[i] else raw_headers[i].upper()
            col_to_cat[i] = cat
            if cat not in unique_cats:
                unique_cats.append(cat)

        all_chunks = []
        for cat in unique_cats:
            # 挑出 Date (0) 與 屬於該 cat 的索引
            sub_indices = [0] + [i for i, c in col_to_cat.items() if c == cat]
            
            # 建立標題
            sub_headers = ["Date"]
            for i in sub_indices:
                if i == 0: continue
                clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "")
                sub_headers.append(clean_h)
            
            # 提取數據
            df_temp = data.iloc[:, sub_indices].copy()
            df_temp.columns = sub_headers
            
            # 插入 B, C 欄 (模擬手作版：Media 與 Product)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 強制轉換日期 (對第一欄 Date 做轉換)
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 強制轉換數值 (對第 4 欄以後做轉換)
            for col in df_temp.columns[3:]:
                # 確保傳入的是 Series (df_temp[col])
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            all_chunks.append(df_temp)

        # 合併
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)
        # 最後補一次零，確保沒有 NaN 導致建模程式崩潰
        df_media_final = df_media_final.fillna(0)

        st.success("✅ 數據處理完畢！")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx")
        with c2:
            st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx")

    except Exception as e:
        st.error(f"❌ 錯誤詳細資訊：{e}")
