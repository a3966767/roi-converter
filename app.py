import streamlit as st
import pandas as pd
from io import BytesIO

# 設定置中佈局
st.set_page_config(page_title="ROI 數據轉換工具 v2.8", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("支援功能：xlsx/csv 雙格式上傳、跨欄標籤辨識、排除 Non-Media、自動補零。")

# 修改 type 參數，同時接受 xlsx 與 csv
uploaded_file = st.file_uploader("第一步：選擇您的檔案 (Excel 或 CSV)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # --- 判斷檔案格式並讀取 ---
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'csv':
            # CSV 讀取，不設 header
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            # Excel 讀取，不設 header
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 處理第 2 行跨欄置中標籤 (向後填充) ---
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        
        # --- 處理第 4 行標題 ---
        headers = df_raw.iloc[3, :].astype(str)

        # --- 處理數據主體 ---
        data = df_raw.iloc[4:].copy()
        
        # 2. 定義索引分類
        first_col_idx = 0
        business_col_indices = [first_col_idx]
        media_col_indices = [first_col_idx]

        for i in range(1, len(groups)):
            group_name = str(groups.iloc[i]).upper()
            
            # 判斷 A: 包含 KPI -> Business
            if "KPI" in group_name:
                business_col_indices.append(i)
            # 判斷 B: 包含 MEDIA 但排除 NON MEDIA -> Media
            elif "MEDIA" in group_name and "NON MEDIA" not in group_name:
                media_col_indices.append(i)

        # 3. 提取數據與處理重複標題
        def handle_duplicates(cols):
            counts = {}
            new_cols = []
            for col in cols:
                if col in counts:
                    counts[col] += 1
                    new_cols.append(f"{col}_{counts[col]}")
                else:
                    counts[col] = 0
                    new_cols.append(col)
            return new_cols

        df_business = data.iloc[:, business_col_indices]
        df_business.columns = handle_duplicates(headers.iloc[business_col_indices].tolist())
        
        df_media = data.iloc[:, media_col_indices]
        df_media.columns = handle_duplicates(headers.iloc[media_col_indices].tolist())

        # --- 數據清理：日期格式與補零 ---
        try:
            date_col = df_business.columns[0]
            df_business[date_col] = pd.to_datetime(df_business[date_col]).dt.date
            df_media[date_col] = pd.to_datetime(df_media[date_col]).dt.date
        except:
            pass

        df_business = df_business.fillna(0)
        df_media = df_media.fillna(0)

        # --- 顯示預覽 (放在上方) ---
        st.success("✅ 檔案解析完成，請確認預覽數據：")
        
        st.subheader("📁 Business 預覽")
        st.dataframe(df_business.head(10), use_container_width=True)

        st.subheader("📁 Media 預覽")
        st.dataframe(df_media.head(10), use_container_width=True)

        st.divider() # 分隔線

        # --- 下載區 (放在最下方) ---
        st.subheader("第二步：點擊按鈕下載檔案")
        
        def to_excel(df):
            output = BytesIO()
            # 下載統一轉為 Excel 格式 (較好閱讀與後續使用)
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="💾 下載 ROI_Business.xlsx",
                data=to_excel(df_business),
                file_name="ROI_Business.xlsx",
                key="dl_biz",
                use_container_width=True
            )
        with col_dl2:
            st.download_button(
                label="💾 下載 ROI_Media.xlsx",
                data=to_excel(df_media),
                file_name="ROI_Media.xlsx",
                key="dl_media",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"❌ 處理失敗，請檢查檔案內容是否正確：{e}")
