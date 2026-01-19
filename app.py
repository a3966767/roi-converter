import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ROI 數據轉換工具 v4.7", layout="centered")

st.title("📊 ROI 數據自動分類轉換器")
st.info("規則更新：區分一般/特殊媒體、標題規範化、排序優化，解決轉換報錯。")

uploaded_file = st.file_uploader("選擇原始 Excel/CSV", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 1. 讀取數據
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df_raw = pd.read_csv(uploaded_file, header=None) if file_ext == 'csv' else pd.read_excel(uploaded_file, header=None)
        
        # 處理第 2 行跨欄標籤 (ffill) 與 第 4 行標題
        groups = df_raw.iloc[1, :].fillna(method='ffill')
        raw_headers = df_raw.iloc[3, :].astype(str).str.strip().tolist()
        data = df_raw.iloc[4:].copy()

        # --- Business 處理 ---
        biz_idx = [0] + [i for i, g in enumerate(groups) if "KPI" in str(g).upper() and i > 0]
        df_biz = data.iloc[:, biz_idx].copy()
        df_biz.columns = ["Date"] + [raw_headers[i] for i in biz_idx if i > 0]
        # 強制日期字串化，解決 effect_start 對齊問題
        df_biz["Date"] = pd.to_datetime(df_biz["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
        df_biz = df_biz.fillna(0)

        # --- Media 處理 ---
        media_idx = [0] + [i for i, g in enumerate(groups) if "MEDIA" in str(g).upper() and i > 0]
        special_keywords = ["OWNED MEDIA", "SHARED MEDIA", "EARNED MEDIA"]
        
        normal_cats = []   # 一般媒體桶
        special_cats = []  # 特殊媒體桶
        col_info = {}

        # 掃描所有媒體欄位並套用規則 1 & 2
        for i in media_idx:
            if i == 0: continue
            g_name = str(groups[i]).upper()
            h_name = raw_headers[i]
            parts = h_name.split("_")
            
            if any(k in g_name for k in special_keywords):
                # 規則 2: 特殊媒體 -> 取第一個底線前
                cat = parts[0].upper()
                if cat not in special_cats: special_cats.append(cat)
                col_info[i] = {"cat": cat, "type": "special"}
            else:
                # 規則 1: 一般媒體 -> 取最後一個底線前
                cat = "_".join(parts[:-1]).upper() if len(parts) > 1 else h_name.upper()
                if cat not in normal_cats: normal_cats.append(cat)
                col_info[i] = {"cat": cat, "type": "normal"}

        # 排序：一般在前，特殊在後 (規則 4)
        sorted_cats = normal_cats + [c for c in special_cats if c not in normal_cats]
        
        all_chunks = []
        for cat in sorted_cats:
            # 找到屬於此分類的原始索引
            target_indices = [i for i, info in col_info.items() if info["cat"] == cat]
            if not target_indices: continue
            
            sub_indices = [0] # 永遠包含日期
            sub_headers = ["Date"]
            
            for i in target_indices:
                sub_indices.append(i)
                h_name = raw_headers[i]
                parts = h_name.split("_")
                
                # 規則 3: 僅留最後底線後文字，首字大寫
                raw_sub_h = parts[-1] if len(parts) > 1 else h_name
                clean_h = raw_sub_h.capitalize()
                
                # 自動修正常見拼寫
                if clean_h == "Spent": clean_h = "Spend"
                sub_headers.append(clean_h)
            
            # 建立子 DataFrame
            df_temp = data.iloc[:, sub_indices].copy()
            df_temp.columns = sub_headers
            
            # 插入必要欄位 (對齊手作版索引)
            df_temp.insert(1, 'Media', cat)
            df_temp.insert(2, 'Product', 'illuma')
            
            # 【關鍵】類型轉換防護 (解決 arg must be a list 報錯)
            # 1. 日期轉換
            df_temp["Date"] = pd.to_datetime(df_temp["Date"], errors='coerce').dt.strftime('%Y-%m-%d')
            
            # 2. 數值轉換 (僅針對 C 欄以後的欄位逐一處理)
            for col_name in df_temp.columns[3:]:
                # 使用 pd.Series 明確轉換每一欄
                df_temp[col_name] = pd.to_numeric(df_temp[col_name], errors='coerce').fillna(0.0)
                
            all_chunks.append(df_temp)

        # 合併所有數據塊
        df_media_final = pd.concat(all_chunks, axis=0, ignore_index=True)

        st.success("✅ 規則處理成功！")

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
