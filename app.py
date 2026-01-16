import streamlit as st
import pandas as pd
from io import BytesIO

# 設定置中佈局
st.set_page_config(page_title="ROI 數據轉換工具 v3.0", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("功能：維持寬表格格式，並在 B 欄自動提取底線前的媒體分類名稱。")

uploaded_file = st.file_uploader("第一步：選擇您的檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # --- 1. 讀取檔案 ---
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        # 處理跨欄標籤與標題
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        headers = df_raw.iloc[3, :].astype(str)
        data = df_raw.iloc[4:].copy()

        # --- 2. 分類索引 ---
        first_col_idx = 0
        biz_idx = [first_col_idx]
        media_idx = [first_col_idx]

        for i in range(1, len(groups)):
            g_name = str(groups[i]).upper()
            if "KPI" in g_name:
                biz_idx.append(i)
            elif "MEDIA" in g_name and "NON MEDIA" not in g_name:
                media_idx.append(i)

        # --- 3. 處理 ROI_Business (維持原樣) ---
        df_biz = data.iloc[:, biz_idx].copy()
        def handle_dupes(cols):
            counts = {}
            new = []
            for c in cols:
                if c in counts:
                    counts[c] += 1
                    new.append(f"{c}_{counts[c]}")
                else:
                    counts[c] = 0
                    new.append(c)
            return new
        df_biz.columns = handle_dupes(headers.iloc[biz_idx].tolist())
        df_biz = df_biz.fillna(0)

        # --- 4. 處理 ROI_Media (新增 B 欄並延伸日期) ---
        # 提取媒體分類 (底線前文字)
        media_headers = headers.iloc[media_idx].tolist()
        categories = []
        for h in media_headers:
            categories.append(h.split("_")[0] if "_" in h else h)

        # 找出所有不重複的媒體分類 (排除第一欄 Date)
        unique_cats = []
        for cat in categories[1:]:
            if cat not in unique_cats:
                unique_cats.append(cat)

        final_media_list = []
        date_header = media_headers[0]

        # 核心邏輯：針對每個「媒體分類」建立獨立的數據區塊
        for cat in unique_cats:
            # 找出屬於該分類的欄位索引
            this_cat_indices = [media_idx[0]] # 必定包含 Date
            this_cat_headers = [date_header]
            
            for i in range(1, len(categories)):
                if categories[i] == cat:
                    this_cat_indices.append(media_idx[i])
                    this_cat_headers.append(media_headers[i])
            
            # 提取該媒體分類的數據
            df_temp = data.iloc[:, this_cat_indices].copy()
            df_temp.columns = this_cat_headers
            
            # 在 B 欄插入分類名稱
            df_temp.insert(1, 'Media', cat)
            final_media_list.append(df_temp)

        # 合併所有媒體區塊 (日期會因此往下重複延伸)
        df_media_final = pd.concat(final_media_list, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0)

        # --- 日期純化 ---
        try:
            date_col = df_biz.columns[0]
            df_biz[date_col] = pd.to_datetime(df_biz[date_col]).dt.date
            df_media_final[df_media_final.columns[0]] = pd.to_datetime(df_media_final[df_media_final.columns[0]]).dt.date
        except:
            pass

        # --- 介面顯示 ---
        st.success("✅ 處理完成！ROI_Media 已新增 B 欄媒體分類。")
        st.subheader("📁 Business 預覽")
        st.dataframe(df_biz.head(5), use_container_width=True)
        st.subheader("📁 Media 預覽")
        st.dataframe(df_media_final.head(10), use_container_width=True)

        st.divider()
        st.subheader("第二步：點擊按鈕下載檔案")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("💾 下載 ROI_Business.xlsx", to_excel(df_biz), "ROI_Business.xlsx", key="dl_biz", use_container_width=True)
        with col_dl2:
            st.download_button("💾 下載 ROI_Media.xlsx", to_excel(df_media_final), "ROI_Media.xlsx", key="dl_media", use_container_width=True)

    except Exception as e:
        st.error(f"❌ 處理失敗：{e}")
