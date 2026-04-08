import streamlit as st
import pandas as pd
import numpy as np
import io
import re  # ★ 新增：用於方法二的 list comprehension

# ==========================================
# 0. 密碼保護機制
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("🔒 請輸入 AE 部門共用密碼以啟用工具：", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("🔒 請輸入 AE 部門共用密碼以啟用工具：", type="password", on_change=password_entered, key="password")
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ==========================================
# 1. 共通資料清理函數
# ==========================================
def clean_dpci(series):
    if series is None:
        return series
    cleaned = series.astype(str)
    cleaned = cleaned.str.replace(r'\s+', '', regex=True)
    cleaned = cleaned.str.replace(r'[/\\]', '-', regex=True)
    cleaned = cleaned.str.replace(r'\.0$', '', regex=True)
    return cleaned

def clean_upc(series):
    if series is None:
        return series
    cleaned = series.astype(str)
    cleaned = cleaned.str.replace(r'\.0$', '', regex=True)
    cleaned = cleaned.str.replace(r'\s+', '', regex=True)
    cleaned = cleaned.replace('nan', np.nan)
    return cleaned

# ==========================================
# 2. 定義資料處理函數 (模組化)
# ==========================================
def process_standard_po(df):
    # ★ 修正：改用 list comprehension 取代 PyArrow 不相容的 str.replace
    df.columns = [re.sub(r'[\ufeff\n\r]', '', str(col)).strip() for col in df.columns]
    po_col = next((c for c in df.columns if 'PO NUMBER' in c.upper()), None)
    if not po_col:
        if next((c for c in df.columns if 'PO' in c.upper() and '#' in c.upper()), None):
            st.error("❌ 您上傳的似乎是【現代版 PO】，請切換到「📈 現代版」分頁！")
        else:
            st.error("❌ 找不到 'PO NUMBER' 欄位！請確認上傳正確的【標準版 PO】檔案。")
        st.stop()

    df[po_col] = df[po_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df = df[df[po_col].str.match(r'^\d+$', na=False)].copy()
    df['PO NUMBER'] = df[po_col]
    df['ASSORTMENT ITEM?'] = df['ASSORTMENT ITEM?'].fillna('N').astype(str).str.strip().str.upper()

    for col in ['DEPARTMENT', 'CLASS', 'ITEM', 'COMPONENT DEPARTMENT', 'COMPONENT CLASS', 'COMPONENT ITEM']:
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

    df['PO UPC'] = clean_upc(df.get('ITEM BAR CODE', pd.Series(np.nan)))
    return df

def process_modern_po(df):
    # ★ 修正：改用 list comprehension 取代 PyArrow 不相容的 str.replace
    df.columns = [re.sub(r'[\ufeff\n\r]', '', str(col)).strip() for col in df.columns]
    po_col = next((c for c in df.columns if 'PO' in c.upper() and '#' in c.upper()), None)
    cost_col = (next((c for c in df.columns if c.upper() == 'COST $'), None) or
                next((c for c in df.columns if 'REV COST' in c.upper() and '$' in c.upper()), None) or
                next((c for c in df.columns if 'COST' in c.upper() and '$' in c.upper()), None))
    retail_col = (next((c for c in df.columns if c.upper() == 'RETAIL $'), None) or
                  next((c for c in df.columns if 'REV RETAIL' in c.upper() and '$' in c.upper()), None) or
                  next((c for c in df.columns if 'RETAIL' in c.upper() and '$' in c.upper()), None))

    if not po_col or not cost_col:
        if next((c for c in df.columns if 'PO NUMBER' in c.upper()), None):
            st.error("❌ 您上傳的似乎是【標準版 PO】，請切換到「📊 標準版」分頁！")
        else:
            st.error("❌ 找不到 'PO #' 或 'COST $' 相關欄位！請確認上傳正確的【現代版 PO】檔案。")
        st.stop()

    df[po_col] = df[po_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df = df[df[po_col].str.match(r'^\d+$', na=False)].copy()

    orig_qty_col = next((c for c in df.columns if 'ORIGINAL QUANTITY' in c.upper()), None)
    rev_qty_col  = next((c for c in df.columns if 'REVISED QUANTITY' in c.upper()), None)

    for col in [orig_qty_col, rev_qty_col, cost_col, retail_col, 'VCP QUANTITY']:
        if col and col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df['PO NUMBER']            = df[po_col].astype(str)
    df['Original_DPCI']        = clean_dpci(df['DPCI'])
    df['Final_DPCI']           = df['Original_DPCI']
    df['ITEM UNIT COST']       = (df[cost_col]   / df[orig_qty_col]) if orig_qty_col and cost_col   else np.nan
    df['ITEM UNIT RETAIL']     = (df[retail_col] / df[orig_qty_col]) if orig_qty_col and retail_col else np.nan
    df['Final_QTY']            = df[rev_qty_col] if rev_qty_col else (df[orig_qty_col] if orig_qty_col else np.nan)
    df['REVISED QUANTITY']     = df['Final_QTY']
    df['ASSORTMENT ITEM?']     = 'N'
    df['COMPONENT ASSORT QTY'] = np.nan
    df['PO UPC']               = clean_upc(df.get('UPC', pd.Series(np.nan)))
    return df

def process_products(files):
    df_list = [pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f) for f in files]
    if not df_list:
        return pd.DataFrame()
    master = pd.concat(df_list, ignore_index=True)
    if 'DPCI' in master.columns:
        master['DPCI'] = clean_dpci(master['DPCI'])
    master['Target UPC'] = clean_upc(master['Barcode'] if 'Barcode' in master.columns else master.get('UPC', pd.Series(np.nan)))

    for col in ['FCA Factory City Unit Cost', 'FOB Unit Cost', 'Suggested Unit Retail', 'Case Unit Quantity', 'Ent Ttl Rcpt U']:
        if col in master.columns:
            master[col] = pd.to_numeric(master[col], errors='coerce')

    if 'FCA Factory City Unit Cost' in master.columns and 'FOB Unit Cost' in master.columns:
        master['Final_Product_Cost'] = master['FCA Factory City Unit Cost'].fillna(master['FOB Unit Cost'])
    elif 'FCA Factory City Unit Cost' in master.columns:
        master['Final_Product_Cost'] = master['FCA Factory City Unit Cost']
    elif 'FOB Unit Cost' in master.columns:
        master['Final_Product_Cost'] = master['FOB Unit Cost']
    else:
        master['Final_Product_Cost'] = np.nan
    return master

def process_assortments(files):
    df_list = []
    for f in files:
        raw_dfs = ([pd.read_csv(f, header=None)] if f.name.endswith('.csv')
                   else [pd.ExcelFile(f).parse(s, header=None) for s in pd.ExcelFile(f).sheet_names])
        for raw_df in raw_dfs:
            header_idx = next(
                (i for i, row in raw_df.iterrows()
                 if row.astype(str).str.replace(r'\s+', '', regex=True).str.lower().str.contains('assortmentdpci', na=False).any()),
                -1
            )
            if header_idx == -1:
                continue
            df = raw_df.iloc[header_idx + 1:].reset_index(drop=True)
            # ★ 修正：改用 list comprehension 取代 PyArrow 不相容的 str.replace
            df.columns = [re.sub(r'[\n\r]', ' ', str(col)).strip() for col in raw_df.iloc[header_idx]]
            cols = {
                'master': next((c for c in df.columns if 'assortment dpci' in c.lower()), None),
                'sub':    next((c for c in df.columns if 'component item dpci' in c.lower() or 'item dpci' in c.lower()), None),
                'cost':   next((c for c in df.columns if 'asst cost' in c.lower() or 'fa box cost' in c.lower()), None),
                'units':  next((c for c in df.columns if 'units in assortment' in c.lower()), None),
            }
            if not all(cols.values()):
                continue
            temp = df[[cols['master'], cols['sub'], cols['cost'], cols['units']]].copy()
            temp.columns = ['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment']
            temp['Assortment_DPCI'] = temp['Assortment_DPCI'].replace(r'^\s*$', np.nan, regex=True).ffill()
            temp['Asst_Box_Cost']   = temp['Asst_Box_Cost'].replace(r'^\s*$', np.nan, regex=True).ffill()
            temp = temp.dropna(subset=['Assortment_DPCI', 'Component_DPCI'])
            temp['Assortment_DPCI'] = clean_dpci(temp['Assortment_DPCI'])
            temp['Component_DPCI']  = clean_dpci(temp['Component_DPCI'])
            temp = temp[~temp['Assortment_DPCI'].str.lower().str.contains('iafillsout|nan|none', na=False)]
            temp['Asst_Box_Cost']       = pd.to_numeric(temp['Asst_Box_Cost'], errors='coerce')
            temp['Units_in_Assortment'] = pd.to_numeric(temp['Units_in_Assortment'], errors='coerce')
            df_list.append(temp)

    if df_list:
        return pd.concat(df_list, ignore_index=True).drop_duplicates(subset=['Assortment_DPCI', 'Component_DPCI'])
    return pd.DataFrame(columns=['Assortment_DPCI', 'Component_DPCI', 'Asst_Box_Cost', 'Units_in_Assortment'])


# ==========================================
# ★ 改進 #2：共用比對引擎（消除重複邏輯）
# ==========================================
def run_validation(po_df, prod_df, asst_files, mode='standard'):
    """
    通用比對引擎，標準版與現代版共用。
    mode: 'standard' | 'modern'
    """
    prod_subset = prod_df[
        [c for c in ['DPCI', 'Final_Product_Cost', 'Suggested Unit Retail',
                     'Case Unit Quantity', 'Ent Ttl Rcpt U', 'Target UPC']
         if c in prod_df.columns]
    ].drop_duplicates(subset=['DPCI'])

    merged = pd.merge(po_df, prod_subset, left_on='Final_DPCI', right_on='DPCI', how='left')

    # 混裝箱邏輯
    if asst_files:
        asst_df = process_assortments(asst_files)
        if mode == 'standard':
            merged = pd.merge(merged, asst_df,
                              left_on=['Original_DPCI', 'Final_DPCI'],
                              right_on=['Assortment_DPCI', 'Component_DPCI'], how='left')
            merged['Target_Cost'] = np.where(
                merged['ASSORTMENT ITEM?'] == 'Y', merged['Asst_Box_Cost'], merged['Final_Product_Cost'])
        else:
            condensed = asst_df.groupby('Assortment_DPCI', as_index=False).agg({'Asst_Box_Cost': 'first'})
            merged = pd.merge(merged, condensed, left_on='Original_DPCI', right_on='Assortment_DPCI', how='left')
            merged['ASSORTMENT ITEM?'] = np.where(merged['Asst_Box_Cost'].notna(), 'Y', 'N')
            merged['Target_Cost'] = np.where(
                merged['ASSORTMENT ITEM?'] == 'Y', merged['Asst_Box_Cost'], merged['Final_Product_Cost'])
    else:
        merged['Target_Cost'] = merged['Final_Product_Cost']

    # (1) 成本比對
    merged['Cost Match'] = (
        np.isclose(merged['ITEM UNIT COST'].fillna(-1), merged['Target_Cost'].fillna(-1), atol=0.01) &
        merged['Target_Cost'].notna()
    )

    # (2) 零售價比對
    if mode == 'standard':
        merged['Retail Match'] = np.where(
            merged['ASSORTMENT ITEM?'] == 'Y', True,
            np.isclose(merged['ITEM UNIT RETAIL'].fillna(0),
                       merged.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01)
        )
    else:
        merged['Retail Match'] = np.isclose(
            merged['ITEM UNIT RETAIL'].fillna(0),
            merged.get('Suggested Unit Retail', pd.Series(np.nan)).fillna(0), atol=0.01
        )

    # (3) 裝箱數比對
    if mode == 'standard':
        merged['Target Case / Assort QTY'] = np.where(
            merged['ASSORTMENT ITEM?'] == 'Y',
            merged.get('Units_in_Assortment', pd.Series(np.nan)),
            merged.get('Case Unit Quantity', pd.Series(np.nan))
        )
        merged['PO VCP / Assort QTY'] = np.where(
            merged['ASSORTMENT ITEM?'] == 'Y',
            merged['COMPONENT ASSORT QTY'], merged['VCP QUANTITY']
        )
    else:
        merged['Target Case / Assort QTY'] = merged.get('Case Unit Quantity', pd.Series(np.nan))
        merged['PO VCP / Assort QTY']      = merged['VCP QUANTITY']
        merged['Target Case / Assort QTY'] = np.where(
            merged['ASSORTMENT ITEM?'] == 'Y',
            merged['PO VCP / Assort QTY'], merged['Target Case / Assort QTY']
        )

    merged['Case QTY Match'] = (
        np.isclose(merged['PO VCP / Assort QTY'].fillna(-1),
                   merged['Target Case / Assort QTY'].fillna(-1), atol=0.01) &
        merged['Target Case / Assort QTY'].notna()
    )
    if mode == 'modern':
        merged['Case QTY Match'] = np.where(
            merged['ASSORTMENT ITEM?'] == 'Y', True, merged['Case QTY Match'])

    # (4) 總數量比對
    merged['Target Commit QTY'] = merged.get('Ent Ttl Rcpt U', pd.Series(np.nan))
    merged['PO Total QTY']      = merged.groupby('Final_DPCI')['Final_QTY'].transform('sum')
    merged['Total QTY Match']   = (
        np.isclose(merged['PO Total QTY'].fillna(-1), merged['Target Commit QTY'].fillna(-1), atol=0.01) &
        merged['Target Commit QTY'].notna()
    )

    # ★ 改進 #3：UPC 比對 — 任一方無資料 → N/A → 不算失敗
    upc_both_exist    = merged['PO UPC'].notna() & merged['Target UPC'].notna()
    merged['UPC Match'] = np.where(
        upc_both_exist,
        merged['PO UPC'] == merged['Target UPC'],
        True  # 無法比對 → 不算異常
    )
    merged['UPC Note'] = np.where(upc_both_exist, '', '⚠️ 無 UPC 可比對')

    # 總判定
    merged['All Match (Pass)'] = (
        merged['Cost Match'] & merged['Retail Match'] &
        merged['Case QTY Match'] & merged['Total QTY Match'] & merged['UPC Match']
    )

    display_cols = [
        'PO NUMBER', 'ASSORTMENT ITEM?', 'Original_DPCI', 'Final_DPCI', 'ITEM DESCRIPTION', 'Final_QTY',
        'Cost Match',      'ITEM UNIT COST',      'Target_Cost',
        'Retail Match',    'ITEM UNIT RETAIL',    'Suggested Unit Retail',
        'Case QTY Match',  'PO VCP / Assort QTY', 'Target Case / Assort QTY',
        'Total QTY Match', 'PO Total QTY',         'Target Commit QTY',
        'UPC Match',       'PO UPC',               'Target UPC',  'UPC Note',
        'All Match (Pass)'
    ]
    return merged[[c for c in display_cols if c in merged.columns]]


# ==========================================
# ★ 改進 #4：顏色標示結果表格
# ==========================================
MATCH_COLS = ['Cost Match', 'Retail Match', 'Case QTY Match', 'Total QTY Match', 'UPC Match', 'All Match (Pass)']

def style_results(df):
    """
    比對欄位套用顏色：
    - True（且有實際 UPC 資料）→ 綠色
    - False                    → 紅色
    - True（但無 UPC 可比對）   → 灰色（N/A）
    """
    styles = pd.DataFrame('', index=df.index, columns=df.columns)

    for col in MATCH_COLS:
        if col not in df.columns:
            continue
        for idx in df.index:
            val = df.at[idx, col]
            is_upc_na = (col == 'UPC Match' and
                         'UPC Note' in df.columns and
                         df.at[idx, 'UPC Note'] != '')
            if is_upc_na:
                styles.at[idx, col] = 'background-color: #e0e0e0; color: #666666;'  # 灰色 N/A
            elif val is True or val == True:
                styles.at[idx, col] = 'background-color: #c6efce; color: #276221;'  # 綠色
            elif val is False or val == False:
                styles.at[idx, col] = 'background-color: #ffc7ce; color: #9c0006;'  # 紅色

    return styles

def show_results(result_df, file_prefix):
    """顯示統計摘要、彩色表格，並提供下載。"""
    errors_df = result_df[result_df['All Match (Pass)'] == False]
    total    = len(result_df)
    err_cnt  = len(errors_df)
    pass_cnt = total - err_cnt

    # 統計卡片
    col1, col2, col3 = st.columns(3)
    col1.metric("📋 總筆數", total)
    col2.metric("✅ 通過",  pass_cnt)
    col3.metric("❌ 異常",  err_cnt)

    if err_cnt == 0:
        st.balloons()
        st.success("🎉 所有資料皆一致！")
    else:
        st.warning(f"⚠️ 發現 {err_cnt} 筆異常（整張表格已標示，紅色 = 異常、灰色 = 無 UPC 可比對）")

    # 彩色表格（顯示全部資料，顏色一目了然）
    styled = result_df.style.apply(style_results, axis=None)
    st.dataframe(styled, use_container_width=True)

    # 下載完整報告
    csv_data = result_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        f"📥 下載完整核對報告 ({file_prefix})",
        data=csv_data,
        file_name=f'PO_Validation_{file_prefix}.csv',
        mime='text/csv'
    )


# ==========================================
# 3. 建立 Streamlit 網頁介面
# ==========================================
st.set_page_config(page_title="訂單自動核對系統", layout="wide")
st.title("📦 跨專案訂單自動核對系統")

st.sidebar.header("📂 步驟 1：上傳共通資料庫")
product_files = st.sidebar.file_uploader("上傳 產品資料表 (可多選)", type=['csv', 'xlsx'], accept_multiple_files=True)
asst_files    = st.sidebar.file_uploader("上傳 混裝箱表單 (可多選/選填)", type=['csv', 'xlsx'], accept_multiple_files=True)

tab1, tab2 = st.tabs(["📊 標準版 (Standard PO) 核對", "📈 現代版 (Modern PO) 核對"])

# ---- 分頁 1：標準版 ----
with tab1:
    st.subheader("上傳標準版 PO 並執行核對")
    po_file_std = st.file_uploader("📥 上傳 Purchase Order Item Details (CSV)", type=['csv'], key="std_po")

    if st.button("🚀 開始核對標準版", type="primary", key="btn_std"):
        if not product_files or not po_file_std:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「標準版 PO」！")
        else:
            with st.spinner("標準版資料清洗與比對中..."):
                po_df   = process_standard_po(pd.read_csv(po_file_std))
                prod_df = process_products(product_files)
                result  = run_validation(po_df, prod_df, asst_files, mode='standard')
            show_results(result, 'Standard')

# ---- 分頁 2：現代版 ----
with tab2:
    st.subheader("上傳現代版 PO 並執行核對")
    po_file_mod = st.file_uploader("📥 上傳 Modern PO Visibility (CSV)", type=['csv'], key="mod_po")

    if st.button("🚀 開始核對現代版", type="primary", key="btn_mod"):
        if not product_files or not po_file_mod:
            st.warning("⚠️ 請確保已在側邊欄上傳「產品資料表」，並在上方上傳「現代版 PO」！")
        else:
            with st.spinner("現代版資料清洗與比對中..."):
                po_df   = process_modern_po(pd.read_csv(po_file_mod))
                prod_df = process_products(product_files)
                result  = run_validation(po_df, prod_df, asst_files, mode='modern')
            show_results(result, 'Modern')
