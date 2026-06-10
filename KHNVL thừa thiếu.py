import streamlit as st
import pandas as pd
import io
import re
from functools import reduce
from openpyxl.styles import Alignment, Border, Side, PatternFill

# --- Streamlit App Configuration ---
st.set_page_config(layout="wide", page_title="BOM Allocation Tool")
st.title("BOM Matching & Stock Distribution Tool")
st.markdown("Upload your BOMs, Inventory, and define your production plan to allocate stock.")

# --- Session State Initialization ---
def init_state(key, default):
    if key not in st.session_state: st.session_state[key] = default

init_state('processed_df', None)
init_state('pivot', None)
init_state('pivot_calculated', None)
init_state('merged_inventory', None)
init_state('allocated_df', None)
init_state('product_cols', [])
init_state('product_priorities', {})

# --- Helper Functions ---
def load_excel_header_search(uploaded_file, sheet_keyword, keywords, is_bom=False, bom_type='RDBOM'):
    if uploaded_file is None: return None
    content_io = io.BytesIO(uploaded_file.getvalue())
    file_extension = uploaded_file.name.lower().split('.')[-1]
    engine_to_use = 'openpyxl'

    if file_extension == 'xls':
        try:
            import xlrd
            if xlrd.__version__ >= '2.0.1':
                engine_to_use = 'xlrd'
        except ImportError:
            pass

    kwargs = {'engine': engine_to_use}
    if engine_to_use == 'xlrd':
        kwargs['engine_kwargs'] = {'ignore_workbook_corruption': True}

    try:
        xls = pd.ExcelFile(content_io, **kwargs)
        sheet_name = next((sn for sn in xls.sheet_names if str(sheet_keyword).lower() in str(sn).lower()), None) if sheet_keyword else xls.sheet_names[0]
        if sheet_name is None: return None

        content_io.seek(0)
        temp_df = pd.read_excel(content_io, sheet_name=sheet_name, header=None, **kwargs)
        header_idx = 0
        for i, row in temp_df.iterrows():
            row_vals = [str(val).strip().lower() for val in row.values]
            if all(kw.lower() in row_vals for kw in keywords):
                header_idx = i
                break

        content_io.seek(0)
        df = pd.read_excel(content_io, sheet_name=sheet_name, header=header_idx, **kwargs)
        if is_bom:
            if bom_type == 'RDBOM': df['Source_RDBOM'] = uploaded_file.name; df['Level'] = df['Level'].astype(str); df = df.ffill()
            if bom_type == 'MANBOM': df['Source_MANBOM'] = uploaded_file.name
        return df
    except Exception as e:
        st.error(f"Error loading {uploaded_file.name}: {str(e)}")
        return None

def process_boms(rdbom_files, manbom_files):
    rdbom_df = pd.concat([load_excel_header_search(f, None, ['Level', 'VNPT P/N'], True, 'RDBOM') for f in rdbom_files], ignore_index=True) if rdbom_files else None
    manbom_df = pd.concat([load_excel_header_search(f, None, ['VNPT P/N', 'Tỉ lệ tiêu hao'], True, 'MANBOM') for f in manbom_files], ignore_index=True) if manbom_files else None

    if rdbom_df is None: return None, None

    rdbom = rdbom_df.copy()
    rdbom['Base_Project'] = rdbom['Source_RDBOM'].str.replace(r'\s*\(\d+\)', '', regex=True).str.replace(r'\.xls[x]?$', '', regex=True, flags=re.IGNORECASE)
    if manbom_df is not None:
        manbom = manbom_df.copy()
        manbom['Base_Project'] = manbom['Source_MANBOM'].str.replace(r'\s*\(\d+\)', '', regex=True).str.replace(r'_ManBom', '', regex=True, flags=re.IGNORECASE).str.replace(r'\.xls[x]?$', '', regex=True, flags=re.IGNORECASE)
        manbom = manbom.drop_duplicates(subset=['VNPT P/N', 'Base_Project'])
        merged = pd.merge(rdbom, manbom, on=['VNPT P/N', 'Base_Project'], how='left', suffixes=('', '_MANBOM'))
    else:
        merged = rdbom

    merged['Source_RDBOM'] = merged['Base_Project']
    if 'Tỉ lệ tiêu hao' in merged.columns: merged = merged.rename(columns={'Tỉ lệ tiêu hao': 'consumption rate'})
    for col in ['Quantity / Product ', 'Quantity / Product']:
        if col in merged.columns: merged = merged.rename(columns={col: 'Quantity/Product'})
    merged['consumption rate'] = pd.to_numeric(merged.get('consumption rate', 0), errors='coerce').fillna(0)
    merged['Quantity/Product'] = pd.to_numeric(merged.get('Quantity/Product', 0), errors='coerce').fillna(0)
    merged['Standard quantity'] = merged['Quantity/Product'] + merged['consumption rate']

    processed = merged.copy()
    is_dup = processed.duplicated(subset=['Source_RDBOM', 'Level', 'VNPT MAN P/N'], keep='first')
    processed['Filter VNPT MAN P/N'] = processed['VNPT MAN P/N'].fillna("")
    processed.loc[is_dup, 'Filter VNPT MAN P/N'] = ""
    valid = processed[processed['Filter VNPT MAN P/N'] != ""]
    counts = valid['Filter VNPT MAN P/N'].value_counts()
    processed['Popularity'] = processed['Filter VNPT MAN P/N'].map(counts)

    g_dict = {k: " | ".join([f"{r['Level']},{r['Source_RDBOM']}" for _, r in v.iterrows()]) for k, v in valid.groupby('Filter VNPT MAN P/N')}
    processed['Level Group'] = processed['Filter VNPT MAN P/N'].map(g_dict)

    processed['Pop_Num'] = pd.to_numeric(processed['Popularity'], errors='coerce')
    idx_max = processed.groupby(['Level', 'Source_RDBOM'])['Pop_Num'].transform('idxmax')
    mask = (processed['Filter VNPT MAN P/N'] == "") & idx_max.notna()
    processed.loc[mask, 'Level Group'] = processed.loc[idx_max[mask], 'Level Group'].values
    processed = processed[processed['Filter VNPT MAN P/N'] != ""].drop(columns=['Pop_Num'])

    if 'Description' not in processed.columns: processed['Description'] = ''

    desired_cols = ['Source_RDBOM', 'Source_MANBOM', 'VNPT P/N', 'Level', 'Description', 'VNPT MAN P/N', 'Quantity/Product', 'consumption rate', 'Standard quantity', 'Filter VNPT MAN P/N', 'Popularity', 'Level Group']
    final_cols = [col for col in desired_cols if col in processed.columns]
    processed = processed[final_cols]

    pivot = pd.pivot_table(processed, index=["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity"], columns=["Source_RDBOM"], values="Standard quantity", aggfunc="sum", fill_value=0).reset_index()
    return processed, pivot.sort_values(by="Level Group").reset_index(drop=True)

def process_inventory(f_tot, f_clc, f_tech, f_scbh, f_khhv):
    dfs = []
    
    # 1. kho_tot
    df_tot = load_excel_header_search(f_tot, "Nhập xuất tồn", ["mã vật tư", "mô tả"])
    if df_tot is not None and not df_tot.empty:
        if str(df_tot.iloc[0].values[0]).strip().startswith('('): df_tot = df_tot.iloc[1:].reset_index(drop=True)
        cuoi_ky_cols = [col for col in df_tot.columns if 'cuối kỳ' in str(col).lower()]
        target_cols = [col for col in df_tot.columns if str(col).lower() in ['mã vật tư']] + cuoi_ky_cols
        if len(target_cols) >= 2:
            df_tot = df_tot[target_cols].copy()
            df_tot.columns = ['VNPT Man P/N', 'Tồn kho tốt'] + list(df_tot.columns[2:])
            dfs.append(df_tot[['VNPT Man P/N', 'Tồn kho tốt']])
            
    # 2. kho_clc
    df_clc = load_excel_header_search(f_clc, "Chi tiết tốt", ["mã vật tư", "mô tả"])
    if df_clc is not None and not df_clc.empty:
        if str(df_clc.iloc[0].values[0]).strip().startswith('('): df_clc = df_clc.iloc[1:].reset_index(drop=True)
        target_cols = [col for col in df_clc.columns if str(col).lower() in ['mã vật tư', 'tồn cuối kỳ']]
        if len(target_cols) >= 2:
            df_clc = df_clc[target_cols].copy()
            df_clc.columns = ['VNPT Man P/N', 'Tồn kho clc']
            dfs.append(df_clc)

    # 3. nha_may_tech
    df_tech = load_excel_header_search(f_tech, "TECH TỔNG", ["row labels", "tên vật tư"])
    if df_tech is not None and not df_tech.empty:
        target_cols = [col for col in df_tech.columns if str(col).lower() in ['row labels', 'tồn cuối']]
        if len(target_cols) >= 2:
            df_tech = df_tech[target_cols].copy()
            df_tech.columns = ['VNPT Man P/N', 'Tồn NM tech']
            dfs.append(df_tech)

    # 4. nha_may_scbh
    df_scbh = load_excel_header_search(f_scbh, "SCBH", ["row labels", "mã vật tư"])
    if df_scbh is not None and not df_scbh.empty:
        target_cols = [col for col in df_scbh.columns if str(col).lower() in ['mã vật tư', 'tồn cuối']]
        if 'Row Labels' in df_scbh.columns and not any(str(c).lower() == 'mã vật tư' for c in target_cols):
            target_cols.insert(0, 'Row Labels')
        if len(target_cols) >= 2:
            df_scbh = df_scbh[target_cols].copy()
            df_scbh.columns = ['VNPT Man P/N', 'Tồn NM scbh']
            dfs.append(df_scbh)

    # 5. khhv
    df_khhv = load_excel_header_search(f_khhv, "TH", ["vnpt pn", "description"])
    if df_khhv is not None and not df_khhv.empty:
        tong_cols = [col for col in df_khhv.columns if 'tổng' in str(col).lower()]
        last_tong = [tong_cols[-1]] if tong_cols else []
        base_cols = [col for col in df_khhv.columns if str(col).lower() in ['vnpt p/n', 'vnpt pn']]
        target_cols = base_cols + last_tong
        if len(target_cols) >= 2:
            df_khhv = df_khhv[target_cols].copy()
            df_khhv.columns = ['VNPT Man P/N', 'Tồn KHHV']
            dfs.append(df_khhv)
            
    if dfs:
        for i in range(len(dfs)):
            dfs[i]['VNPT Man P/N'] = dfs[i]['VNPT Man P/N'].astype(str).str.strip()
            dfs[i] = dfs[i][~dfs[i]['VNPT Man P/N'].str.lower().isin(['', 'nan', '(blank)', 'none', 'null'])]

        merged_inventory = reduce(lambda left, right: pd.merge(left, right, on='VNPT Man P/N', how='outer'), dfs)
        stock_cols = [col for col in merged_inventory.columns if col != 'VNPT Man P/N']
        for col in stock_cols:
            merged_inventory[col] = pd.to_numeric(merged_inventory[col], errors='coerce').fillna(0)

        sum_cols = [col for col in stock_cols if col != 'Tồn kho clc']
        merged_inventory['Tổng tồn'] = merged_inventory[sum_cols].sum(axis=1)
        return merged_inventory
    return None

# --- UI Workflow ---
with st.expander("1. Upload BOM Files & Process", expanded=True):
    rdbom_files = st.file_uploader("Upload RDBOM.xlsx (Required)", type=["xlsx", "xls"], accept_multiple_files=True)
    manbom_files = st.file_uploader("Upload MANBOM.xlsx (Optional)", type=["xlsx", "xls"], accept_multiple_files=True)
    if st.button("Process BOMs"):
        with st.spinner("Processing..."):
            processed, pivot = process_boms(rdbom_files, manbom_files)
            if pivot is not None:
                st.session_state.processed_df, st.session_state.pivot = processed, pivot
                st.success("BOMs processed successfully!")

                st.subheader("Final BOM Result Table")
                st.dataframe(st.session_state.processed_df)

                st.subheader("Pivot Table")
                st.dataframe(st.session_state.pivot)

with st.expander("2. Production Plan & Calculate Demand", expanded=True):
    if st.session_state.pivot is not None:
        st.markdown("### Enter Production Quantity (KHSX) for each product")
        pivot_df = st.session_state.pivot.copy()

        index_cols = ["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity"]
        rdbom_cols = [col for col in pivot_df.columns if col not in index_cols]

        cols = st.columns(3)
        multipliers = {}

        for i, col in enumerate(rdbom_cols):
            with cols[i % 3]:
                multipliers[col] = st.number_input(f"Multiplier for {col}", min_value=0.0, value=1.0, step=1.0)

        if st.button("Calculate Demand"):
            for col in rdbom_cols:
                pivot_df[f"{col} - Calculated"] = pivot_df[col] * multipliers[col]

            st.session_state.pivot_calculated = pivot_df
            st.success("Component quantities calculated successfully!")
            st.dataframe(st.session_state.pivot_calculated)
    else:
        st.info("Please complete Step 1 to generate the pivot table first.")

with st.expander("3. Upload Inventory Files", expanded=True):
    if st.session_state.pivot_calculated is not None:
        col1, col2 = st.columns(2)
        with col1:
            f_tot = st.file_uploader("Upload 'Kho Tốt'", type=["xlsx", "xls"])
            f_clc = st.file_uploader("Upload 'Kho CLC'", type=["xlsx", "xls"])
            f_tech = st.file_uploader("Upload 'Nhà máy Tech'", type=["xlsx", "xls"])
        with col2:
            f_scbh = st.file_uploader("Upload 'Nhà máy SCBH'", type=["xlsx", "xls"])
            f_khhv = st.file_uploader("Upload 'KHHV'", type=["xlsx", "xls"])
            
        if st.button("Process Inventory"):
            with st.spinner("Processing Inventory..."):
                merged_inv = process_inventory(f_tot, f_clc, f_tech, f_scbh, f_khhv)
                if merged_inv is not None:
                    st.session_state.merged_inventory = merged_inv
                    
                    # Merge into Pivot
                    pivot_final = st.session_state.pivot_calculated.copy()
                    cols_to_merge = [c for c in ['VNPT Man P/N', 'Tồn kho tốt', 'Tồn kho clc', 'Tồn NM tech', 'Tồn NM scbh', 'Tồn KHHV', 'Tổng tồn'] if c in merged_inv.columns]
                    
                    pivot_final = pd.merge(
                        pivot_final,
                        merged_inv[cols_to_merge],
                        left_on='Filter VNPT MAN P/N',
                        right_on='VNPT Man P/N',
                        how='left'
                    )
                    if 'VNPT Man P/N' in pivot_final.columns:
                        pivot_final = pivot_final.drop(columns=['VNPT Man P/N'])
                        
                    st.session_state.pivot_calculated = pivot_final
                    st.success("Inventory processed and joined successfully!")
                    st.dataframe(st.session_state.pivot_calculated)
                else:
                    st.warning("Could not extract inventory data. Please check the uploaded files.")
    else:
        st.info("Please calculate demand in Step 2 before uploading inventory.")
