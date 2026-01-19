import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.4", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正：明確指定 Series 進行類型轉換，解決 arg must be a list 報錯。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV 檔案", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據 (不轉換型態)
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        # 處理第 2 行跨欄標籤 (ffill) 與 第 4 行標題
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- 2. 處理 Business ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        
        # 強制轉換 Date 欄位 (指定 Series)
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 Media ---
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and "NON MEDIA" not in str(g).upper() and i > 0]

        # 找出分類
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
                    clean_h = raw_headers[i].replace(f"{cat.lower()}_", "").replace(f"{cat.upper()}_", "").strip()
                    sub_headers.append(clean_h)
            
            # 建立子表
            df_temp = data.iloc[:, sub_indices].copy()
            
            # 處理重複標題
            final_h, h_cnt = [], {}
            for h in sub_headers:
                if h in h_cnt:
                    h_cnt[h] += 1
                    final_h.append(f"{h}_{h_cnt[h]}")
                else:
                    h_cnt[h] = 0
                    final_h.append(h)
            df_temp.columns = final_h
            
            # 插入 B, C 欄 (模擬手作版索引)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 【關鍵修正：類型轉換】
            # 使用明確的欄位選取方式，確保傳入 pd.to_datetime/pd.to_numeric 的是 Series
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
            
            for col in df_temp.columns[3:]: # 從第 4 欄開始是數值
                # 這裡確保對 Series 執行轉換
                df_temp[col] = pd.to_numeric(df_temp[col], errors='coerce').fillna(0.0)
            
            all_chunks.append(df_temp)

        # 合併所有區塊
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0)

        st.success("✅ 轉換成功！類型報錯已修正。")

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
