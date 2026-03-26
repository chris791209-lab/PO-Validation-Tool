import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 0. 密碼保護機制
# ==========================================
def check_password():
    """回傳 True 代表使用者輸入了正確的密碼"""
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    
    elif not st.session_state["password_correct"]:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==========================================
# 1. 共通 DPCI 清理函數
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
# 2. 定義資料處理函數 (模組化)
# ==========================================
def process_standard_po(df):
    """處理【標準版】訂單原始資料"""
    df.columns = df.columns.str.replace('\ufeff', '').str.strip()
    if 'PO NUMBER' not in df.columns:
        st.error("❌ 檔案讀取錯誤：找不到 'PO NUMBER' 欄位！請確認您上傳的是【標準版】。")
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

    qty_series = np.where(df['ASSORTMENT ITEM?'] == 'Y', df['COMPONENT ITEM TOTAL QTY'], df['TOTAL ITEM QTY'])
    df['Final_QTY'] = pd.Series(qty_series).astype(str).str.replace(',', '', regex=False).astype(float)
    
    for col in ['ITEM UNIT COST', 'ITEM UNIT RETAIL', 'VCP QUANTITY', 'COMPONENT ASSORT QTY']:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def process_modern_po(df):
    """處理【現代版】訂單原始資料"""
    df.columns = df.columns.str.replace('\ufeff', '').str.strip()
    if 'PO #' not in df.columns or 'COST $' not in df.columns:
        st.error("❌ 檔案讀取錯誤：找不到 'PO #' 或 'COST $' 欄位！請確認您上傳的是【現代版】。")
        st.stop()
        
    df = df[df['PO #'].astype(str).str.match(r'^\d+$', na=False)].copy()
    
    # 【新增】將 REVISED QUANTITY 也納入數字清理轉換
    for col in ['ORIGINAL QUANTITY', 'REVISED QUANTITY', 'COST $', 'RETAIL $', 'VCP QUANTITY']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    df['PO NUMBER'] = df['PO #'].astype(str)
    df['Original_DPCI'] = clean_dpci(df['DPCI'])
    df['Final_DPCI'] = df['Original_DPCI']
    df['Final_QTY'] = df['ORIGINAL QUANTITY']
    
    df['ITEM UNIT COST'] = df['COST $'] / df['ORIGINAL QUANTITY']
    df['ITEM UNIT RETAIL'] = df['RETAIL $'] / df['ORIGINAL QUANTITY']
    df['ASSORTMENT ITEM?'] = 'N'
    df['COMPONENT ASSORT QTY'] = np.nan
    return df

def process_products(files):
    df_list = []
    for f in files:
        df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
        df_list.append(df)
    if not df_list: return pd.DataFrame()
        
    master_product_df = pd.concat(df_list, ignore_index=True)
    if 'DPCI' in master_product_df.columns:
        master_product_df['DPCI'] = clean_dpci(master_product_df['DPCI'])
    
    # 【新增】將 Ent Ttl Rcpt U 納入數值讀取欄位
    numeric_cols = ['FCA Factory City Unit Cost', 'FOB Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Ent Ttl Rcpt U']
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
    df_list = []
    for f in files:
        raw_dfs = [pd.read_csv(f, header=None)] if f.name.endswith('.csv') else [pd.ExcelFile(f).parse(sheet_name, header=None) for sheet_name in pd.ExcelFile(f).sheet_names]
        for raw_df in raw_dfs:
            header_idx = next((i for i, row in raw_df.iterrows() if row.astype(str).str.replace(r'\s+', '', regex=True).str.lower().str.contains('assortmentdpci', na=False).any()), -1)
            if header_idx != -1:
                df = raw_df.iloc[header_idx + 1:].reset_index(drop=True)
                df.columns = raw_df.iloc[header_idx].astype(str).str.replace(r'[\n\r]', ' ', regex=True).str.strip()
                
                cols = {
                    'master': next((c for c in df.columns if 'assortment dpci' in c.lower() or 'assortmentdpci' in c.replace(' ', '').lower()), None),
                    'sub': next((c for c in df.columns if 'component item dpci' in c.lower() or 'item dpci' in c.lower()), None),
                    'cost': next((c for c in df.columns if 'asst cost' in c.lower() or 'fa box cost' in c.lower()), None),
                    'units': next((c for c in df.columns if 'units in assortment' in c.lower()), None)
                }
                
                if all(cols.values()):
                    temp_df = df[[cols['master'], cols['sub'], cols['cost'], cols['units']]].copy()
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
        return pd.concat(df_list, ignore_index=True).drop_duplicates(subset=['Assortment_DPCI', 'Component_DPCI'])
    return pd.DataFrame(columns=['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment'])

# ==========================================
# 3. 建立 Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("📦 跨專案訂單自動核對系統")

st.sidebar.header("📂 步驟 1：上傳共通資料庫")
product_files = st.sidebar.file_uploader("上傳 產品資料表 (可多選)", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files = st.sidebar.file_uploader("上傳 混裝箱表單 (可多選/選填)", type=['csv', 'xlsx'], accept_multiple_files=True)

tab1, tab2 = st.tabs(["📊 標準版 (Standard PO) 核對", "📈 現代版 (Modern PO) 核對"])

# ----------------- 分頁 1: 標準版 -----------------
with tab1:
    st.subheader("上傳標準版 PO 並執行核對")
    po_file_std = st.file_uploader("📥 上傳 Purchase Order Item Details (CSV)", type=['csv'], key="std_po")
    
    if st.button("🚀 開始核對標準版", type="primary", key="btn_std"):
        if not product_files or not po_file_std:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「標準版 PO」！")
        else:
            with st.spinner("標準版資料清洗與比對中..."):
                po_df = process_standard_po(pd.read_csv(po_file_std))
                prod_df = process_products(product_files)
                
                # 【新增】抓取 Ent Ttl Rcpt U 欄位
                prod_subset = prod_df[[c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Ent Ttl Rcpt U'] if c in prod_df.columns]].drop_duplicates(subset=['DPCI'])
                merged_df = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

                if asst_files:
                    asst_df = process_assortments(asst_files)
                    merged_df = pd.merge(merged_df, asst_df, left_on=['Original_DPCI', 'Final_DPCI'], right_on=['Assortment_DPCI', 'Component_DPCI'], how='left')
                    merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
                else:
                    merged_df['Target_Cost'] = merged_df['Final_Product_Cost']

                # (1) 成本比對
                merged_df['Cost Match'] = np.isclose(merged_df['ITEM UNIT COST'].fillna(-1), merged_df['Target_Cost'].fillna(-1), atol=0.01)
                merged_df['Cost Match'] = np.where(merged_df['Target_Cost'].isna(), False, merged_df['Cost Match'])

                # (2) 零售價比對
                merged_df['Retail Match'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', True, np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01))
                
                # (3) 裝箱數比對
                merged_df['Target Case / Assort QTY'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df.get('Units_in_Assortment', pd.Series(np.nan)), merged_df.get('Case Unit Quantity', pd.Series(np.nan)))
                merged_df['PO VCP / Assort QTY'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['COMPONENT ASSORT QTY'], merged_df['VCP QUANTITY'])
                merged_df['Case QTY Match'] = np.isclose(merged_df['PO VCP / Assort QTY'].fillna(-1), merged_df['Target Case / Assort QTY'].fillna(-1), atol=0.01)
                merged_df['Case QTY Match'] = np.where(merged_df['Target Case / Assort QTY'].isna(), False, merged_df['Case QTY Match'])
                
                # 【新增】(4) 總數量比對 (標準版使用 Final_QTY 也就是 TOTAL ITEM QTY/COMPONENT ITEM TOTAL QTY)
                merged_df['Target Total QTY'] = merged_df.get('Ent Ttl Rcpt U', pd.Series(np.nan))
                merged_df['PO Total QTY'] = merged_df['Final_QTY']
                merged_df['Total QTY Match'] = np.isclose(merged_df['PO Total QTY'].fillna(-1), merged_df['Target Total QTY'].fillna(-1), atol=0.01)
                merged_df['Total QTY Match'] = np.where(merged_df['Target Total QTY'].isna(), False, merged_df['Total QTY Match'])
                
                merged_df['All Match (Pass)'] = merged_df['Cost Match'] & merged_df['Retail Match'] & merged_df['Case QTY Match'] & merged_df['Total QTY Match']
                
                # 顯示結果
                display_cols = [
                    'PO NUMBER', 'ASSORTMENT ITEM?', 'Original_DPCI', 'Final_DPCI', 'ITEM DESCRIPTION', 
                    'Cost Match', 'ITEM UNIT COST', 'Target_Cost', 
                    'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail', 
                    'Case QTY Match', 'PO VCP / Assort QTY', 'Target Case / Assort QTY', 
                    'Total QTY Match', 'PO Total QTY', 'Target Total QTY',
                    'All Match (Pass)'
                ]
                result_df = merged_df[[c for c in display_cols if c in merged_df.columns]]
                errors_df = result_df[result_df['All Match (Pass)'] == False]
                
                st.success(f"✅ 核對完成！總共 {len(result_df)} 筆，其中發現 {len(errors_df)} 筆異常。")
                if len(errors_df) > 0: st.dataframe(errors_df)
                else: st.balloons(); st.info("🎉 所有資料皆一致！")
                
                csv_data = result_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載標準版核對報告", data=csv_data, file_name='PO_Validation_Standard.csv', mime='text/csv')

# ----------------- 分頁 2: 現代版 -----------------
with tab2:
    st.subheader("上傳現代版 PO 並執行核對")
    po_file_mod = st.file_uploader("📥 上傳 Modern PO Visibility (CSV)", type=['csv'], key="mod_po")
    
    if st.button("🚀 開始核對現代版", type="primary", key="btn_mod"):
        if not product_files or not po_file_mod:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「現代版 PO」！")
        else:
            with st.spinner("現代版資料清洗與比對中..."):
                po_df = process_modern_po(pd.read_csv(po_file_mod))
                prod_df = process_products(product_files)
                
                # 【新增】抓取 Ent Ttl Rcpt U 欄位
                prod_subset = prod_df[[c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Ent Ttl Rcpt U'] if c in prod_df.columns]].drop_duplicates(subset=['DPCI'])
                merged_df = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

                if asst_files:
                    asst_df = process_assortments(asst_files)
                    condensed_asst = asst_df.groupby('Assortment_DPCI', as_index=False).agg({'Asst_Box_Cost': 'first'})
                    merged_df = pd.merge(merged_df, condensed_asst, left_on='Original_DPCI', right_on='Assortment_DPCI', how='left')
                    merged_df['ASSORTMENT ITEM?'] = np.where(merged_df['Asst_Box_Cost'].notna(), 'Y', 'N')
                    merged_df['Target_Cost'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['Asst_Box_Cost'], merged_df['Final_Product_Cost'])
                else:
                    merged_df['Target_Cost'] = merged_df['Final_Product_Cost']

                # (1) 成本比對
                merged_df['Cost Match'] = np.isclose(merged_df['ITEM UNIT COST'].fillna(-1), merged_df['Target_Cost'].fillna(-1), atol=0.01)
                merged_df['Cost Match'] = np.where(merged_df['Target_Cost'].isna(), False, merged_df['Cost Match'])

                # (2) 零售價比對
                merged_df['Retail Match'] = np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01)
                
                # (3) 裝箱數比對
                merged_df['Target Case / Assort QTY'] = merged_df.get('Case Unit Quantity', pd.Series(np.nan))
                merged_df['PO VCP / Assort QTY'] = merged_df['VCP QUANTITY']
                merged_df['Target Case / Assort QTY'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', merged_df['PO VCP / Assort QTY'], merged_df['Target Case / Assort QTY'])
                
                merged_df['Case QTY Match'] = np.isclose(merged_df['PO VCP / Assort QTY'].fillna(-1), merged_df['Target Case / Assort QTY'].fillna(-1), atol=0.01)
                merged_df['Case QTY Match'] = np.where(merged_df['Target Case / Assort QTY'].isna(), False, merged_df['Case QTY Match'])
                merged_df['Case QTY Match'] = np.where(merged_df['ASSORTMENT ITEM?'] == 'Y', True, merged_df['Case QTY Match'])
                
                # 【新增】(4) 總數量比對 (現代版特別抓取 REVISED QUANTITY 進行比對)
                merged_df['Target Total QTY'] = merged_df.get('Ent Ttl Rcpt U', pd.Series(np.nan))
                merged_df['PO Total QTY'] = merged_df['REVISED QUANTITY']
                merged_df['Total QTY Match'] = np.isclose(merged_df['PO Total QTY'].fillna(-1), merged_df['Target Total QTY'].fillna(-1), atol=0.01)
                merged_df['Total QTY Match'] = np.where(merged_df['Target Total QTY'].isna(), False, merged_df['Total QTY Match'])
                
                merged_df['All Match (Pass)'] = merged_df['Cost Match'] & merged_df['Retail Match'] & merged_df['Case QTY Match'] & merged_df['Total QTY Match']
                
                # 顯示結果
                display_cols = [
                    'PO NUMBER', 'ASSORTMENT ITEM?', 'Original_DPCI', 'Final_DPCI', 'ITEM DESCRIPTION', 
                    'Cost Match', 'ITEM UNIT COST', 'Target_Cost', 
                    'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail', 
                    'Case QTY Match', 'PO VCP / Assort QTY', 'Target Case / Assort QTY', 
                    'Total QTY Match', 'PO Total QTY', 'Target Total QTY',
                    'All Match (Pass)'
                ]
                result_df = merged_df[[c for c in display_cols if c in merged_df.columns]]
                errors_df = result_df[result_df['All Match (Pass)'] == False]
                
                st.success(f"✅ 核對完成！總共 {len(result_df)} 筆，其中發現 {len(errors_df)} 筆異常。")
                if len(errors_df) > 0: st.dataframe(errors_df)
                else: st.balloons(); st.info("🎉 所有資料皆一致！")
                
                csv_data = result_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載現代版核對報告", data=csv_data, file_name='PO_Validation_Modern.csv', mime='text/csv')
