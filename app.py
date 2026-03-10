import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. 定義資料處理函數
# ==========================================
def process_po(df):
    """處理訂單(PO)原始資料，生成最終核對用的 DPCI 與數量"""
    # 確保所有欄位名稱沒有前後空白
    df.columns = df.columns.str.strip()
    
    # 填補空值，避免後續文字處理報錯
    df['ASSORTMENT ITEM?'] = df['ASSORTMENT ITEM?'].fillna('N').astype(str).str.strip().str.upper()
    
    # 處理可能帶有小數點的浮點數轉字串問題 (如 240.0 -> "240")
    cols_to_clean = ['DEPARTMENT', 'CLASS', 'ITEM', 'COMPONENT DEPARTMENT', 'COMPONENT CLASS', 'COMPONENT ITEM']
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    # 1. 判斷並合成 DPCI
    # 如果是 Y，合成 COMPONENT 的 DPCI；否則合成原始 DPCI
    df['Final_DPCI'] = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        df['COMPONENT DEPARTMENT'].str.zfill(3) + "-" + df['COMPONENT CLASS'].str.zfill(2) + "-" + df['COMPONENT ITEM'].str.zfill(4),
        df['DEPARTMENT'].str.zfill(3) + "-" + df['CLASS'].str.zfill(2) + "-" + df['ITEM'].str.zfill(4)
    )

    # 2. 判斷並抓取最終數量
    # 如果是 Y，抓 COMPONENT ITEM TOTAL QTY；否則抓 TOTAL ITEM QTY
    qty_series = np.where(
        df['ASSORTMENT ITEM?'] == 'Y',
        df['COMPONENT ITEM TOTAL QTY'],
        df['TOTAL ITEM QTY']
    )
    
    # 將數量欄位中的千分位逗號去掉，並轉為數字
    df['Final_QTY'] = pd.Series(qty_series).astype(str).str.replace(',', '', regex=False).astype(float)
    
    # 金額轉換確保為數字
    df['ITEM UNIT COST'] = pd.to_numeric(df['ITEM UNIT COST'], errors='coerce')
    df['ITEM UNIT RETAIL'] = pd.to_numeric(df['ITEM UNIT RETAIL'], errors='coerce')
    df['VCP QUANTITY'] = pd.to_numeric(df['VCP QUANTITY'], errors='coerce')
    
    return df

def process_products(files):
    """處理並合併多份產品資料庫"""
    df_list = []
    for f in files:
        if f.name.endswith('.csv'):
            df = pd.read_csv(f)
        else:
            df = pd.read_excel(f)
        df_list.append(df)
    
    # 將多份產品表縱向合併成一份大表
    master_product_df = pd.concat(df_list, ignore_index=True)
    
    # 確保 DPCI 欄位存在且為標準字串格式 (避免 "240-11-1234 " 這類帶空白的狀況)
    if 'DPCI' in master_product_df.columns:
        master_product_df['DPCI'] = master_product_df['DPCI'].astype(str).str.strip()
    
    # 確保需要的數值欄位為數字
    numeric_cols = ['FCA Factory City Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity']
    for col in numeric_cols:
        if col in master_product_df.columns:
            master_product_df[col] = pd.to_numeric(master_product_df[col], errors='coerce')
            
    return master_product_df

# ==========================================
# 2. 建立 Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("🎃 Halloween 訂單自動核對系統")
st.markdown("只需上傳客人的 **PO 原始資料** 與 **產品資料總表**，系統會自動展開混裝箱(Assortment)並進行核對。")

# 側邊欄：檔案上傳區
st.sidebar.header("📂 檔案上傳區")
po_file = st.sidebar.file_uploader("1. 上傳 PO 原始資料 (CSV)", type=['csv'])
product_files = st.sidebar.file_uploader("2. 上傳產品資料表 (可多選, Excel/CSV)", type=['csv', 'xlsx'], accept_multiple_files=True)

# 主畫面操作邏輯
if po_file and product_files:
    if st.button("🚀 開始核對訂單", type="primary"):
        with st.spinner("資料處理中，請稍候..."):
            
            # 讀取並處理資料
            po_df = process_po(pd.read_csv(po_file))
            prod_df = process_products(product_files)
            
            # 使用 Left Join 將產品資料帶入 PO 表格 (利用我們處理好的 Final_DPCI 作為橋樑)
            # 由於我們只要比對特定欄位，可以先對 prod_df 篩選，避免欄位過多
            prod_subset = prod_df[['DPCI', 'FCA Factory City Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Vendor Product Description *']]
            # 去除產品表中重複的 DPCI，以第一筆為主
            prod_subset = prod_subset.drop_duplicates(subset=['DPCI'])
            
            merged_df = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

            # ==========================================
            # 3. 執行檢驗邏輯
            # ==========================================
            # (1) 成本比對：PO金額 vs 產品表 FOB Cost
            merged_df['Cost Match'] = np.isclose(merged_df['ITEM UNIT COST'].fillna(0), merged_df['FCA Factory City Unit Cost'].fillna(0), atol=0.01)
            
            # (2) 零售價比對：PO零售價 vs 產品表 Suggested Retail
            merged_df['Retail Match'] = np.isclose(merged_df['ITEM UNIT RETAIL'].fillna(0), merged_df['Suggested Unit Retail'].fillna(0), atol=0.01)
            
            # (3) 裝箱數比對：PO VCP vs 產品表 Case Qty
            merged_df['Case QTY Match'] = merged_df['VCP QUANTITY'] == merged_df['Case Unit Quantity']
            
            # 建立一個總覽欄位：只要有一個 False 就是異常
            merged_df['All Match (Pass)'] = merged_df['Cost Match'] & merged_df['Retail Match'] & merged_df['Case QTY Match']
            
            # 整理要顯示/輸出的欄位順序
            display_cols = [
                'PO NUMBER', 'ASSORTMENT ITEM?', 'Final_DPCI', 'ITEM DESCRIPTION', 'Final_QTY',
                'Cost Match', 'ITEM UNIT COST', 'FCA Factory City Unit Cost',
                'Retail Match', 'ITEM UNIT RETAIL', 'Suggested Unit Retail',
                'Case QTY Match', 'VCP QUANTITY', 'Case Unit Quantity', 'All Match (Pass)'
            ]
            
            result_df = merged_df[display_cols]

            # ==========================================
            # 4. 顯示結果並提供下載
            # ==========================================
            errors_df = result_df[result_df['All Match (Pass)'] == False]
            
            st.success(f"✅ 核對完成！總共 {len(result_df)} 筆資料，其中發現 {len(errors_df)} 筆異常。")
            
            st.subheader("⚠️ 異常資料列表 (有不一致的項目)")
            if len(errors_df) > 0:
                st.dataframe(errors_df.style.highlight_max(axis=0))
            else:
                st.info("太棒了！所有資料皆一致，沒有異常。")
                
            st.subheader("📋 完整核對結果")
            st.dataframe(result_df)

            # 將結果轉為 CSV 並提供下載按鈕
            csv_buffer = io.StringIO()
            result_df.to_csv(csv_buffer, index=False)
            # 使用 utf-8-sig 讓 Excel 開啟時不會出現中文亂碼
            csv_data = csv_buffer.getvalue().encode('utf-8-sig')
            
            st.download_button(
                label="📥 下載完整核對報告 (CSV)",
                data=csv_data,
                file_name='PO_Validation_Result.csv',
                mime='text/csv',
            )
else:
    st.info("請先在左側欄上傳 **PO原始資料** 以及至少一份 **產品資料表**。")