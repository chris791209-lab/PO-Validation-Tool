import streamlit as st
def check_password():
    """回傳 True 代表使用者輸入了正確的密碼"""

    def password_entered():
        """檢查使用者輸入的密碼是否與 Streamlit Secrets 中的密碼相符"""
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            # 密碼正確後，刪除 session_state 中的密碼紀錄以策安全
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # 第一次進入網頁，顯示密碼輸入框
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    
    elif not st.session_state["password_correct"]:
        # 密碼輸入錯誤，顯示錯誤訊息並重新要求輸入
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    
    else:
        # 密碼正確
        return True
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. 共通 DPCI 清理函數 (核彈級清理，確保全專案適用)
# ==========================================
def clean_dpci(series):
    """清理 DPCI 字串，強制移除所有空白、斜線與隱藏字元"""
    if series is None:
        return series
    cleaned = series.astype(str)
    cleaned = cleaned.str.replace(r'\s+', '', regex=True)
    cleaned = cleaned.str.replace(r'[/\\]', '-', regex=True)
    cleaned = cleaned.str.replace(r'\.0$', '', regex=True)
    return cleaned

# ==========================================
# 2. 定義資料處理函數 (分為標準版與現代版)
# ==========================================
def process_standard_po(df):
    """處理【標準版】訂單(Purchase Order Item Details)原始資料"""
    df.columns = df.columns.str.replace('\ufeff', '').str.strip()
    
    if 'PO NUMBER' not in df.columns:
        st.error("❌ 檔案讀取錯誤：找不到 'PO NUMBER' 欄位！請確認您上傳的是【標準版】或切換左側的格式選項。")
        st.stop()
        
    df = df[df['PO NUMBER'].astype(str).str.match(r'^\d+$', na=False)].copy()
    df['ASSORTMENT ITEM?'] = df['ASSORTMENT ITEM?'].fillna('N').astype(str).str.strip().str.upper()
    
    cols_to_clean = ['DEPARTMENT', 'CLASS', 'ITEM', 'COMPONENT DEPARTMENT', 'COMPONENT CLASS', 'COMPONENT ITEM']
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    df['Original_DPCI'] = clean_dpci(df['DEPARTMENT'].str.zfill(3) + "-" + df['CLASS'].str.zfill(2) + "-" + df['ITEM'].str.zfill(4))
    
    df['Final_DPCI'] = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        clean_dpci(df['COMPONENT DEPARTMENT'].str.zfill(3) + "-" + df['COMPONENT CLASS'].str.zfill(2) + "-" + df['COMPONENT ITEM'].str.zfill(4)),
        df['Original_DPCI']
    )

    qty_series = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        df['COMPONENT ITEM TOTAL QTY'],
        df['TOTAL ITEM QTY']
    )
    df['Final_QTY'] = pd.Series(qty_series).astype(str).str.replace(',', '', regex=False).astype(float)
    
    df['ITEM UNIT COST'] = pd.to_numeric(df['ITEM UNIT COST'], errors='coerce')
    df['ITEM UNIT RETAIL'] = pd.to_numeric(df['ITEM UNIT RETAIL'], errors='coerce')
    df['VCP QUANTITY'] = pd.to_numeric(df['VCP QUANTITY'], errors='coerce')
    df['COMPONENT ASSORT QTY'] = pd.to_numeric(df['COMPONENT ASSORT QTY'], errors='coerce')
    
    return df

def process_modern_po(df):
    """處理【現代版】訂單(Modern PO Visibility)原始資料"""
    df.columns = df.columns.str.replace('\ufeff', '').str.strip()
    
    if 'PO #' not in df.columns or 'COST $' not in df.columns:
        st.error("❌ 檔案讀取錯誤：找不到 'PO #' 或 'COST $' 欄位！請確認您上傳的是【現代版】或切換左側的格式選項。")
        st.stop()
        
    df = df[df['PO #'].astype(str).str.match(r'^\d+$', na=False)].copy()
    
    # 清理數字欄位 (去除千分位逗號)
    for col in ['ORIGINAL QUANTITY', 'COST $', 'RETAIL $', 'VCP QUANTITY']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    df['PO NUMBER'] = df['PO #'].astype(str)
    
    # 現代版直接使用 DPCI 欄位
    df['Original_DPCI'] = clean_dpci(df['DPCI'])
    df['Final_DPCI'] = df['Original_DPCI']
    df['Final_QTY'] = df['ORIGINAL QUANTITY']
    
    # 【計算單價】：總金額 / ORIGINAL QUANTITY
    df['ITEM UNIT COST'] = df['COST $'] / df['ORIGINAL QUANTITY']
    df['ITEM UNIT RETAIL'] = df['RETAIL $'] / df['ORIGINAL QUANTITY']
    
    # 現代版沒有子單號，預設給予 N，後續透過混裝表判定
    df['ASSORTMENT ITEM?'] = 'N'
    df['COMPONENT ASSORT QTY'] = np.nan
    
    return df

def process_products(files):
    """處理並合併多份產品資料庫"""
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_list.append(df)
    
    if not df_list:
        return pd.DataFrame()
        
    master_product_df = pd.concat(df_list, ignore_index=True)
    
    if 'DPCI' in master_product_df.columns:
        master_product_df['DPCI'] = clean_dpci(master_product_df['DPCI'])
    
    numeric_cols = ['FCA Factory City Unit Cost', 'FOB Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity']
    for col in numeric_cols:
        if col in master_product_df.columns:
            master_product_df[col] = pd.to_numeric(master_product_df[col], errors='coerce')
            
    if 'FCA Factory City Unit Cost' in master_product_df.columns and 'FOB Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FCA Factory City Unit Cost'].fillna(master_product_df['FOB Unit Cost'])
    elif 'FCA Factory City Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FCA Factory City Unit Cost']
    elif 'FOB Unit Cost' in master_product_df.columns:
        master_product_df['Final_Product_Cost'] = master_product_df['FOB Unit Cost']
    else:
        master_product_df['Final_Product_Cost'] = np.nan
            
    return master_product_df

def process_assortments(files):
    """掃描 Excel 所有分頁，並透過智慧模糊比對抓取欄位"""
    df_list = []
    for f in files:
        if f.name.endswith('.csv'):
            raw_dfs = [pd.read_csv(f, header=None)]
        else:
            xl = pd.ExcelFile(f)
            raw_dfs = [xl.parse(sheet_name, header=None) for sheet_name in xl.sheet_names]
            
        for raw_df in raw_dfs:
            header_idx = -1
            for i, row in raw_df.iterrows():
                row_str = row.astype(str).str.replace(r'\s+', '', regex=True).str.lower()
                if row_str.str.contains('assortmentdpci', na=False).any():
                    header_idx = i
                    break
                    
            if header_idx != -1:
                df = raw_df.copy()
                df.columns = df.iloc[header_idx]
                df.columns = df.columns.astype(str).str.replace(r'[\n\r]', ' ', regex=True).str.strip()
                df = df.iloc[header_idx + 1:].reset_index(drop=True)
                
                master_dpci_col = next((c for c in df.columns if 'assortment dpci' in c.lower() or 'assortmentdpci' in c.replace(' ', '').lower()), None)
                sub_dpci_col = next((c for c in df.columns if 'component item dpci' in c.lower() or 'item dpci' in c.lower()), None)
                cost_col = next((c for c in df.columns if 'asst cost' in c.lower() or 'fa box cost' in c.lower()), None)
                units_col = next((c for c in df.columns if 'units in assortment' in c.lower()), None)
                
                if master_dpci_col and sub_dpci_col and cost_col and units_col:
                    temp_df = df[[master_dpci_col, sub_dpci_col, cost_col, units_col]].copy()
                    temp_df.columns = ['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment']
                    
                    temp_df['Assortment_DPCI'] = temp_df['Assortment_DPCI'].replace(r'^\s*$', np.nan, regex=True).ffill()
                    temp_df['Asst_Box_Cost'] = temp_df['Asst_Box_Cost'].replace(r'^\s*$', np.nan, regex=True).ffill()
                    
                    temp_df = temp_df.dropna(subset=['Assortment_DPCI', 'Component_DPCI'])
                    
                    temp_df['Assortment_DPCI'] = clean_dpci(temp_df['Assortment_DPCI'])
                    temp_df['Component_DPCI'] = clean_dpci(temp_df['Component_DPCI'])
                    
                    temp_df = temp_df[~temp_df['Assortment_DPCI'].str.lower().str.contains('iafillsout|nan|none', na=False)]
                    
                    temp_df['Asst_Box_Cost'] = pd.to_numeric(temp_df['Asst_Box_Cost'], errors='coerce')
                    temp_df['Units_in_Assortment'] = pd.to_numeric(temp_df['Units_in_Assortment'], errors='coerce')
                    df_list.append(temp_df)
                
    if df_list:
        master_asst = pd.concat(df_list, ignore_index=True)
        return master_asst.drop_duplicates(subset=['Assortment_DPCI', 'Component_DPCI'])
    return pd.DataFrame(columns=['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment'])

# ==========================================
# 3. 建立 Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("📦 跨專案訂單自動核對系統 (雙版本支援)")

st.sidebar.header("📂 檔案上傳與設定區")
# 【新增】版本切換開關
po_format = st.sidebar.radio(
    "1. 請選擇 PO 原始檔格式:", 
    ["標準版 (Purchase Order Item Details)", "現代版 (Modern PO Visibility)"],
    help="切換不同的 PO 匯出版本，系統將自動套用對應的讀取與計價邏輯。"
)

po_file = st.sidebar.file_uploader("2. 上傳 PO 原始資料 (CSV)", type=['csv'])
product_files = st.sidebar.file_uploader("3. 上傳產品資料表 (可多選, Excel/CSV)", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files = st.sidebar.file_uploader("4. 上傳混裝箱 Assortment 表單 (可多選, 可選填)", type=['csv', 'xlsx'], accept_multiple_files=True)

if po_file and product_files:
    if st.button("🚀 開始核對訂單", type="primary"):
        with st.spinner("資料清洗與比對中，請稍候..."):
            
            # 依據選擇的格式套用不同的 PO 處理函數
            if po_format == "標準版 (Purchase Order Item Details)":
                po_df = process_standard_po(pd.read_csv(po_file))
            else:
                po_df = process_modern_po(pd.read_csv(po_file))
                
            prod_df = process_products(product_files)
            
            available_cols = [c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail', 'Case Unit Quantity'] if c in prod_df.columns]
            prod_subset = prod_df[available_cols].drop_duplicates(subset=['DPCI'])
            
            # 1. 拿子單號與產品主檔 Merge
            merged_df = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

            # 2. 拿母單號+子單號與混裝表 Merge
            if asst_files:
                asst_df = process_assortments(asst_files)
                
                if po_format == "現代版 (Modern PO Visibility)":
                    # 現代版沒有子零件，直接針對母單號取得整箱成本 (避免資料重複展開)
                    condensed_asst = asst_df.groupby('Assortment_DPCI', as_index=False).agg({'Asst_Box_Cost': 'first'})
                    merged_df = pd.merge(merged_df, condensed_asst, left_on='Original_DPCI', right_on='Assortment_DPCI', how='left')
                    
                    # 只要在混裝表找到該單號，就標記為混裝品
                    merged_df['ASSORTMENT ITEM?'] = np.where(merged_df['Asst_Box_Cost'].notna(), 'Y', 'N')
                    merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
                else:
                    # 標準版邏輯
                    merged_df = pd.merge(merged_df, asst_df, 
                                         left_on=['Original_DPCI', 'Final_DPCI'], 
                                         right_on=['Assortment_DPCI', 'Component_DPCI'], 
                                         how='left')
                    merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
            else:
                merged_df['Target_Cost'] = merged_df['Final_Product_Cost']

            # ==========================================
            # 4. 執行檢驗邏輯
            # ==========================================
            # (1) 成本比對
            merged_df['Cost Match'] = np.isclose(merged_df['ITEM UNIT COST'].fillna(-1), merged_df['Target_Cost'].fillna(-1), atol=0.01)
            merged_df['Cost Match'] = np.where(merged_df['Target_Cost'].isna(), False, merged_df['Cost Match'])

            # (2) 零售價比對
            if po_format == "現代版 (Modern PO Visibility)":
                # 現代版已獨立算出單位零售價，直接進行比對
                merged_df['Retail Match'] = np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01)
            else:
                # 標準版的混裝母單號零售價通常不準，直接判定為 True
                merged_df['Retail Match'] = np.where(
                    merged_df['ASSORTMENT ITEM?'] == 'Y',
                    True,
                    np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01)
                )
            
            # (3) 裝箱數比對：動態切換核對欄位
            if po_format == "現代版 (Modern PO Visibility)":
                merged_df['Target Case / Assort QTY'] = merged_df.get('Case Unit Quantity', pd.Series(np.nan))
                merged_df['PO VCP / Assort QTY'] = merged_df['VCP QUANTITY']
            else:
                merged_df['Target Case / Assort QTY'] = np.where(
                    merged_df['ASSORTMENT ITEM?'] == 'Y',
                    merged_df.get('Units_in_Assortment', pd.Series(np.nan)),
                    merged_df.get('Case Unit Quantity', pd.Series(np.nan))
                )
                merged_df['PO VCP / Assort QTY'] = np.where(
                    merged_df['ASSORTMENT ITEM?'] == 'Y',
                    merged_df['COMPONENT ASSORT QTY'],
                    merged_df['VCP QUANTITY']
                )

            # 進行對比
            merged_df['Case QTY Match'] = np.isclose(merged_df['PO VCP / Assort QTY'].fillna(-1), merged_df['Target Case / Assort QTY'].fillna(-1), atol=0.01)
            merged_df['Case QTY Match'] = np.where(merged_df['Target Case / Assort QTY'].isna(), False, merged_df['Case QTY Match'])
            
            # 總覽判定
            merged_df['All Match (Pass)'] = merged_df['Cost Match'] & merged_df['Retail Match'] & merged_df['Case QTY Match']
            
            display_cols = [
                'PO NUMBER', 'ASSORTMENT ITEM?', 'Original_DPCI', 'Final_DPCI', 'ITEM DESCRIPTION', 'Final_QTY',
                'Cost Match', 'ITEM UNIT COST', 'Target_Cost',
                'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail',
                'Case QTY Match', 'PO VCP / Assort QTY', 'Target Case / Assort QTY', 'All Match (Pass)'
            ]
            
            display_cols = [c for c in display_cols if c in merged_df.columns]
            result_df = merged_df[display_cols]

            # ==========================================
            # 5. 顯示結果並提供下載
            # ==========================================
            errors_df = result_df[result_df['All Match (Pass)'] == False]
            
            st.success(f"✅ 核對完成！總共 {len(result_df)} 筆有效訂單，其中發現 {len(errors_df)} 筆異常。")
            
            st.subheader("⚠️ 異常資料列表 (有不一致的項目)")
            if len(errors_df) > 0:
                st.dataframe(errors_df)
            else:
                st.balloons()
                st.info("🎉 太棒了！所有資料皆一致，沒有異常。")
                
            st.subheader("📋 完整核對結果")
            st.dataframe(result_df)

            csv_buffer = io.StringIO()
            result_df.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue().encode('utf-8-sig')
            
            st.download_button(
                label="📥 下載完整核對報告 (CSV)",
                data=csv_data,
                file_name='PO_Validation_Result.csv',
                mime='text/csv',
            )
else:
    st.info("請先在左側欄上傳 **PO原始資料** 以及至少一份 **產品資料表**。")
