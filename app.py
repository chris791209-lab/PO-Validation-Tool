import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. 定義資料處理函數
# ==========================================
def process_po(df):
    """處理訂單(PO)原始資料，生成最終核對用的 DPCI 與數量"""
    df.columns = df.columns.str.strip()
    df['ASSORTMENT ITEM?'] = df['ASSORTMENT ITEM?'].fillna('N').astype(str).str.strip().str.upper()
    
    cols_to_clean = ['DEPARTMENT', 'CLASS', 'ITEM', 'COMPONENT DEPARTMENT', 'COMPONENT CLASS', 'COMPONENT ITEM']
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    # 合成 DPCI
    df['Final_DPCI'] = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        df['COMPONENT DEPARTMENT'].str.zfill(3) + "-" + df['COMPONENT CLASS'].str.zfill(2) + "-" + df['COMPONENT ITEM'].str.zfill(4),
        df['DEPARTMENT'].str.zfill(3) + "-" + df['CLASS'].str.zfill(2) + "-" + df['ITEM'].str.zfill(4)
    )

    # 抓取最終數量
    qty_series = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        df['COMPONENT ITEM TOTAL QTY'],
        df['TOTAL ITEM QTY']
    )
    df['Final_QTY'] = pd.Series(qty_series).astype(str).str.replace(',', '', regex=False).astype(float)
    
    # 確保數值欄位型態
    df['ITEM UNIT COST'] = pd.to_numeric(df['ITEM UNIT COST'], errors='coerce')
    df['ITEM UNIT RETAIL'] = pd.to_numeric(df['ITEM UNIT RETAIL'], errors='coerce')
    df['VCP QUANTITY'] = pd.to_numeric(df['VCP QUANTITY'], errors='coerce')
    
    return df

def process_products(files):
    """處理並合併多份產品資料庫，並解決 FCA/FOB 欄位問題"""
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_list.append(df)
    
    master_product_df = pd.concat(df_list, ignore_index=True)
    
    if 'DPCI' in master_product_df.columns:
        master_product_df['DPCI'] = master_product_df['DPCI'].astype(str).str.strip()
    
    numeric_cols = ['FCA Factory City Unit Cost', 'FOB Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity']
    for col in numeric_cols:
        if col in master_product_df.columns:
            master_product_df[col] = pd.to_numeric(master_product_df[col], errors='coerce')
            
    # 【優化 2】解決 FCA 與 FOB 雙欄位問題：優先取 FCA，若無則取 FOB
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
    """【優化 1】處理混裝表，提取 Component DPCI 與對應的 Item Cost"""
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        
        # 判斷 DPCI 欄位名稱是哪一種
        dpci_col = None
        if 'Component Item DPCI' in df.columns:
            dpci_col = 'Component Item DPCI'
        elif 'Item DPCI' in df.columns:
            dpci_col = 'Item DPCI'
            
        # 抓取價錢欄位
        cost_col = 'Item Cost' if 'Item Cost' in df.columns else None
        
        if dpci_col and cost_col:
            temp_df = df[[dpci_col, cost_col]].rename(columns={dpci_col: 'DPCI', cost_col: 'Asst_Item_Cost'})
            temp_df['DPCI'] = temp_df['DPCI'].astype(str).str.strip()
            temp_df['Asst_Item_Cost'] = pd.to_numeric(temp_df['Asst_Item_Cost'], errors='coerce')
            df_list.append(temp_df)
            
    if df_list:
        master_asst = pd.concat(df_list, ignore_index=True).dropna(subset=['DPCI'])
        return master_asst.drop_duplicates(subset=['DPCI']) # 去除重複
    return pd.DataFrame(columns=['DPCI', 'Asst_Item_Cost'])

# ==========================================
# 2. 建立 Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("Halloween 訂單自動核對系統")

# 側邊欄：檔案上傳區
st.sidebar.header("📂 檔案上傳區")
po_file = st.sidebar.file_uploader("1. 上傳 PO 原始資料 (CSV)", type=['csv'])
product_files = st.sidebar.file_uploader("2. 上傳產品資料表 (可多選, Excel/CSV)", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files = st.sidebar.file_uploader("3. 上傳混裝箱 Assortment 表單 (可多選, 可選填)", type=['csv', 'xlsx'], accept_multiple_files=True)

# 主畫面操作邏輯
if po_file and product_files:
    if st.button("🚀 開始核對訂單", type="primary"):
        with st.spinner("資料處理中，請稍候..."):
            
            po_df = process_po(pd.read_csv(po_file))
            prod_df = process_products(product_files)
            
            # 整理產品主檔的合併欄位
            available_cols = [c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail', 'Case Unit Quantity'] if c in prod_df.columns]
            prod_subset = prod_df[available_cols].drop_duplicates(subset=['DPCI'])
            
            # 1. 先與產品主檔 Merge
            merged_df = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

            # 2. 若有上傳混裝表，與混裝表 Merge 獲取混裝成本
            if asst_files:
                asst_df = process_assortments(asst_files)
                merged_df = pd.merge(merged_df, asst_df, left_on='Final_DPCI', right_on='DPCI', how='left', suffixes=('', '_asst'))
                # 【合併比對標準】：優先使用產品主檔價錢，若為空則使用混裝表價錢
                merged_df['Target_Cost'] = merged_df['Final_Product_Cost'].fillna(merged_df.get('Asst_Item_Cost', np.nan))
            else:
                merged_df['Target_Cost'] = merged_df['Final_Product_Cost']

            # ==========================================
            # 3. 執行檢驗邏輯
            # ==========================================
            # (1) 成本比對：PO金額 vs 目標對照金額 (FCA/FOB/混裝表)
            merged_df['Cost Match'] = np.isclose(merged_df['ITEM UNIT COST'].fillna(0), merged_df['Target_Cost'].fillna(0), atol=0.01)
            
            # (2) 零售價比對
            merged_df['Retail Match'] = np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01)
            
            # (3) 裝箱數比對
            merged_df['Case QTY Match'] = merged_df['VCP QUANTITY'] == merged_df.get('Case Unit Quantity', pd.Series(np.nan))
            
            # 總覽判定
            merged_df['All Match (Pass)'] = merged_df['Cost Match'] & merged_df['Retail Match'] & merged_df['Case QTY Match']
            
            # 整理顯示欄位
            display_cols = [
                'PO NUMBER', 'ASSORTMENT ITEM?', 'Final_DPCI', 'ITEM DESCRIPTION', 'Final_QTY',
                'Cost Match', 'ITEM UNIT COST', 'Target_Cost', # Target_Cost 也就是我們系統判定的標準價格
                'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail',
                'Case QTY Match', 'VCP QUANTITY', 'Case Unit Quantity', 'All Match (Pass)'
            ]
            
            # 確保欄位存在才顯示
            display_cols = [c for c in display_cols if c in merged_df.columns]
            result_df = merged_df[display_cols]

            # ==========================================
            # 4. 顯示結果並提供下載
            # ==========================================
            errors_df = result_df[result_df['All Match (Pass)'] == False]
            
            st.success(f"✅ 核對完成！總共 {len(result_df)} 筆資料，其中發現 {len(errors_df)} 筆異常。")
            
            st.subheader("⚠️ 異常資料列表 (有不一致的項目)")
            if len(errors_df) > 0:
                st.dataframe(errors_df)
            else:
                st.info("太棒了！所有資料皆一致，沒有異常。")
                
            st.subheader("📋 完整核對結果")
            st.dataframe(result_df)

            csv_buffer = io.StringIO()
            result_df.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue().encode('utf-8-sig')
            
            st.download_button(
                label="📥 下載完整核對報告 (CSV)",
                data=csv_data,
                file_name='PO_Validation_Result_V2.csv',
                mime='text/csv',
            )
else:
    st.info("請先在左側欄上傳 **PO原始資料** 以及至少一份 **產品資料表**。")

