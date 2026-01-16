import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v3.1", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("修正：解決重複索引報錯，確保 B 欄媒體分類正確生成。")

uploaded_file = st.file_uploader("第一步：選擇您的檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # --- 1. 讀取檔案 ---
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        # 處理第 2 行跨欄標籤 (ffill)
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        # 處理第 4 行標題
        headers = df_raw.iloc[3, :].astype(str).tolist()
        # 處理數據主體
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

        # 輔助函式：處理重複標題
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

        # --- 3. 處理 ROI_Business ---
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = handle_dupes([headers[i] for i in biz_idx])
        df_biz = df_biz.fillna(0)

        # --- 4. 處理 ROI_Media (新增 B 欄並延伸日期) ---
        # 提取媒體分類 (底線前文字)
        media_categories = []
        for i in media_idx:
            h = headers[i]
            if i == 0: # Date 欄位
                media_categories.append("DATE")
            else:
                media_categories.append(h.split("_")[0] if "_" in h else h)

        # 取得所有不重複的媒體分類
        unique_cats = []
        for cat in media_categories[1:]:
            if cat not in unique_cats:
                unique_cats.append(cat)

        all_media_chunks = []
        
        for cat in unique_cats:
            # 找出屬於該分類的欄位
            cat_sub_indices = [media_idx[0]] # 包含 Date
            cat_sub_headers = [headers[media_idx[0]]]
            
            for i in range(1, len(media_idx)):
                if media_categories[i] == cat:
                    cat_sub_indices.append(media_idx[i])
                    cat_sub_headers.append(headers[media_idx[i]])
            
            # 提取數據
            df_temp = data.iloc[:, cat_sub_indices].copy()
            df_temp.columns = handle_dupes(cat_sub_headers) # 先處理單一區塊的重複
            
            # 在 B 欄位置插入 Media 名稱
            df_temp.insert(1, 'Media', cat)
            all_media_chunks.append(df_temp)

        # 合併所有區塊
        # 使用 ignore_index=True 避免索引重複報錯
        df_media_final = pd.concat(all_media_chunks, axis=0, ignore_index=True)
        df_media_final = df_media_final.fillna(0)

        # --- 數據清理 ---
        try:
            date_col_biz = df_biz.columns[0]
            df_biz[date_col_biz] = pd.to_datetime(df_biz[date_col_biz]).dt.date
            
            date_col_med = df_media_final.columns[0]
            df_media_final[date_col_med] = pd.to_datetime(df_media_final[date_col_med]).dt.date
        except:
            pass

        # --- 介面顯示 ---
        st.success("✅ 處理完成！")
        
        st.subheader("📁 Business 預覽")
        st.dataframe(df_biz.head(5), use_container_width=True)

        st.subheader("📁 Media 預覽 (B 欄為媒體分類)")
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
        st.error(f"❌ 處理失敗，詳細錯誤：{e}")

