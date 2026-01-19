import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.4", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正：解決重複索引報錯 (Reindexing error)，並校準 B 欄與 C 欄格式。")

uploaded_file = st.file_uploader("上傳原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # --- 1. 讀取與基礎處理 ---
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        # 填充跨欄標籤 (Row 2)
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        # 取得標題列 (Row 4)
        raw_headers = df_raw.iloc[3, :].astype(str).tolist()
        # 數據主體 (Row 5+)
        data = df_raw.iloc[4:].copy()

        # 輔助函式：確保標題清單唯一 (避免 Reindexing 錯誤)
        def make_unique(cols):
            counts = {}
            new_cols = []
            for col in cols:
                # 移除前後空白
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
            if "KPI" in str(groups[i]).upper():
                biz_idx.append(i)
        
        df_biz = data.iloc[:, biz_idx].copy()
        # 賦予唯一標題
        biz_header_list = [raw_headers[i] for i in biz_idx]
        biz_header_list[0] = "Date" # 強制第一欄叫 Date
        df_biz.columns = make_unique(biz_header_list)
        df_biz = df_biz.fillna(0)

        # --- 3. 處理 ROI_Media (建立 B 欄 Media 並清理標題) ---
        media_idx = [0]
        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        # 預先計算每個欄位所屬的 Media 分類 (底線前文字)
        col_to_cat = {}
        for i in media_idx:
            if i == 0: 
                col_to_cat[i] = "DATE"
            else:
                h = raw_headers[i]
                col_to_cat[i] = h.split("_")[0].upper() if "_" in h else h.upper()

        # 取得不重複的分類清單
        unique_cats = [c for c in list(dict.fromkeys(col_to_cat.values())) if c != "DATE"]
        
        all_media_chunks = []
        for cat in unique_cats:
            # 挑出屬於該分類的欄位
            cat_sub_idx = [0]
            # 清理標題：移除媒體前綴 (例如移除 "META_" 或 "meta_")
            cat_sub_headers = ["Date"]
            
            for i in media_idx:
                if i == 0: continue
                if col_to_cat[i] == cat:
                    cat_sub_idx.append(i)
                    # 移除前綴邏輯
                    orig_h = raw_headers[i]
                    prefix_lower = f"{cat.lower()}_"
                    prefix_upper = f"{cat.upper()}_"
                    clean_h = orig_h.replace(prefix_lower, "").replace(prefix_upper, "")
                    cat_sub_headers.append(clean_h)
            
            # 建立該媒體的數據區塊
            df_temp = data.iloc[:, cat_sub_idx].copy()
            # **關鍵修正**：在賦予標題前再次確保唯一性，防止 Reindexing 錯誤
            df_temp.columns = make_unique(cat_sub_headers)
            
            # 在 B 欄插入 Media 分類
            df_temp.insert(1, 'Media', cat)
            all_media_chunks.append(df_temp)

        # 合併所有媒體區塊
        # ignore_index=True 確保列索引重新編號，避免衝突
        df_media_final = pd.concat(all_media_chunks, axis=0, ignore_index=True)
        # 這裡不強制補 0，以符合手作版可能存在的空白狀態，或依需求補 0
        df_media_final = df_media_final.fillna(0) 

        # --- 4. 日期格式化 ---
        try:
            df_biz.iloc[:, 0] = pd.to_datetime(df_biz.iloc[:, 0]).dt.date
            df_media_final.iloc[:, 0] = pd.to_datetime(df_media_final.iloc[:, 0]).dt.date
        except:
            pass

        # --- 5. 介面顯示與下載 ---
        st.success("✅ 轉換完成！已修正重複索引問題。")
        
        st.subheader("📁 Business 數據預覽")
        st.dataframe(df_biz.head(5), use_container_width=True)

        st.subheader("📁 Media 數據預覽 (已校準 B/C 欄)")
        st.dataframe(df_media_final.head(10), use_container_width=True)

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.divider()
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx", use_container_width=True)
        with col_dl2:
            st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx", use_container_width=True)

    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")
