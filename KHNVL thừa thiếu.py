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
        merged = pd.merge(rdbom, manbom, on=['VNPT P/N', 'Base_Project'], how='left')
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
    pivot = pd.pivot_table(processed, index=["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity"], columns=["Source_RDBOM"], values="Standard quantity", aggfunc="sum", fill_value=0).reset_index()
    return processed, pivot.sort_values(by="Level Group").reset_index(drop=True)

def process_inventory(f_tot, f_clc, f_tech, f_scbh, f_khhv, pivot):
    inventory_dfs = []
    def get_inv(f, kw_sheet, kws, rename_to, filter_cols): 
        df = load_excel_header_search(f, kw_sheet, kws)
        if df is not None and not df.empty:
            if str(df.iloc[0].values[0]).strip().startswith('('): df = df.iloc[1:].reset_index(drop=True)
            t_cols = []
            for c in filter_cols:
                for col in df.columns:
                    if c in str(col).lower() and col not in t_cols: t_cols.append(col)
            if len(t_cols) >= 2:
                df = df[t_cols[:2]]
                df.columns = ['VNPT Man P/N', rename_to]
                inventory_dfs.append(df)
    
    get_inv(f_tot, "Nhập xuất tồn", ["mã vật tư", "mô tả"], 'Tồn kho tốt', ['mã vật tư', 'cuối kỳ'])
    get_inv(f_clc, "Chi tiết tốt", ["mã vật tư", "mô tả"], 'Tồn kho clc', ['mã vật tư', 'tồn cuối kỳ'])
    get_inv(f_tech, "TECH TỔNG", ["row labels", "tên vật tư"], 'Tồn NM tech', ['row labels', 'tồn cuối'])
    get_inv(f_scbh, "SCBH", ["row labels", "mã vật tư"], 'Tồn NM scbh', ['mã vật tư', 'tồn cuối'])
    get_inv(f_khhv, "TH", ["vnpt pn", "description"], 'Tồn KHHV', ['vnpt', 'tổng'])

    if not inventory_dfs: return pivot
    
    cleaned_dfs = []
    for d in inventory_dfs:
        d['VNPT Man P/N'] = d['VNPT Man P/N'].astype(str).str.strip()
        d = d[~d['VNPT Man P/N'].str.lower().isin(['', 'nan', '(blank)', 'none', 'null'])]
        cleaned_dfs.append(d)

    merged_inv = reduce(lambda l, r: pd.merge(l, r, on='VNPT Man P/N', how='outer'), cleaned_dfs)
    s_cols = [c for c in merged_inv.columns if c != 'VNPT Man P/N']
    for c in s_cols: merged_inv[c] = pd.to_numeric(merged_inv[c], errors='coerce').fillna(0)
    merged_inv['Tổng tồn'] = merged_inv[[c for c in s_cols if c != 'Tồn kho clc']].sum(axis=1)

    cols_to_remove = [col for col in pivot.columns if any(sc in col for sc in ['Tổng tồn', 'TONG_TON', 'Tồn kho tốt', 'Tồn kho clc', 'Tồn NM tech', 'Tồn NM scbh', 'Tồn KHHV'])]
    if cols_to_remove: pivot = pivot.drop(columns=cols_to_remove)
    
    pivot = pd.merge(pivot, merged_inv, left_on='Filter VNPT MAN P/N', right_on='VNPT Man P/N', how='left')
    if 'VNPT Man P/N' in pivot.columns: pivot = pivot.drop(columns=['VNPT Man P/N'])
    return pivot

def run_allocation(pivot, product_cols, multipliers):
    for p in product_cols: pivot[f"{p} - Calculated"] = pivot[p] * multipliers[p]
    parent = {i: i for i in pivot.index}
    def find(i):
        if parent[i] == i: return i
        parent[i] = find(parent[i]); return parent[i]
    def union(i, j): parent[find(i)] = find(j)
    
    for idx, row in pivot.iterrows():
        usages = set([x.strip() for x in str(row['Level Group']).split('|')]) if pd.notna(row['Level Group']) else set()
        for u in usages: 
            indices = [i for i, r in pivot.iterrows() if u in str(r.get('Level Group', ''))]
            for i in range(1, len(indices)): union(indices[0], indices[i])

    pool_dict = {}
    for idx in pivot.index: pool_dict.setdefault(find(idx), []).append(idx)
    
    results = []
    for pool_id, indices in enumerate(pool_dict.values(), 1):
        df_group = pivot.loc[indices].copy()
        df_group['Popularity'] = pd.to_numeric(df_group['Popularity'], errors='coerce')
        stock_col = 'Tổng tồn' if 'Tổng tồn' in df_group.columns else ('TONG_TON' if 'TONG_TON' in df_group.columns else None)
        if not stock_col: df_group['Tổng tồn'] = 0; stock_col = 'Tổng tồn'
        df_group[stock_col] = pd.to_numeric(df_group[stock_col], errors='coerce').fillna(0)
        df_group = df_group.sort_values(by=['Popularity', stock_col], ascending=[True, True])
        df_group['Allocation Pool'] = pool_id
        for p in product_cols:
            df_group.rename(columns={p: f"{p} - Standard Qty", f"{p} - Calculated": f"{p} - SL theo KH"}, inplace=True)
            df_group[f"{p} - SL sau phân bổ kho"] = 0.0
        group_remain = {p: df_group[f"{p} - SL theo KH"].max() if f"{p} - SL theo KH" in df_group.columns else 0.0 for p in product_cols}
        main_idx = df_group.index[0]

        for idx, row in df_group.iterrows():
            stock = row.get(stock_col, 0)
            for p in product_cols:
                if p in str(row.get('Level Group', '')) and group_remain[p] > 0 and stock > 0:
                    use = min(group_remain[p], stock)
                    df_group.at[idx, f"{p} - SL sau phân bổ kho"] = use
                    group_remain[p] -= use; stock -= use
            df_group.at[idx, 'Remaining_Stock'] = stock

        for p in product_cols:
            if group_remain[p] > 0:
                valid_idx = [i for i, r in df_group.iterrows() if p in str(r.get('Level Group', ''))]
                t_idx = valid_idx[0] if valid_idx else main_idx
                df_group.at[t_idx, f"{p} - SL sau phân bổ kho"] += group_remain[p]
                df_group.at[t_idx, 'Remaining_Stock'] -= group_remain[p]
        results.append(df_group)
    return pd.concat(results, ignore_index=True)

def resolve_flaws(allocated_df, product_cols):
    def resolve_pool(df_pool):
        for _ in range(20):
            neg_mask = df_pool['Remaining_Stock'] < -0.0001
            if not neg_mask.any(): break
            fixed = False
            for def_idx in df_pool[neg_mask].index:
                def_val = -df_pool.loc[def_idx, 'Remaining_Stock']
                for p_col in product_cols:
                    alloc_col = f"{p_col} - SL sau phân bổ kho"
                    if alloc_col in df_pool.columns and df_pool.loc[def_idx, alloc_col] > 0:
                        valid_A = [idx for idx, row in df_pool.iterrows() if p_col in str(row.get('Level Group', ''))]
                        for shared_idx in valid_A:
                            if shared_idx == def_idx: continue
                            for p_col_B in product_cols:
                                if p_col_B == p_col: continue
                                alloc_col_B = f"{p_col_B} - SL sau phân bổ kho"
                                if alloc_col_B in df_pool.columns and df_pool.loc[shared_idx, alloc_col_B] > 0:
                                    valid_B = [idx for idx, row in df_pool.iterrows() if p_col_B in str(row.get('Level Group', ''))]
                                    for spare_idx in valid_B:
                                        if spare_idx in [shared_idx, def_idx]: continue
                                        if df_pool.loc[spare_idx, 'Remaining_Stock'] > 0:
                                            shift = min(def_val, df_pool.loc[shared_idx, alloc_col_B], df_pool.loc[spare_idx, 'Remaining_Stock'])
                                            if shift > 0:
                                                df_pool.loc[shared_idx, alloc_col_B] -= shift
                                                df_pool.loc[spare_idx, alloc_col_B] += shift
                                                df_pool.loc[spare_idx, 'Remaining_Stock'] -= shift
                                                df_pool.loc[def_idx, alloc_col] -= shift
                                                df_pool.loc[shared_idx, alloc_col] += shift
                                                df_pool.loc[def_idx, 'Remaining_Stock'] += shift
                                                def_val -= shift
                                                fixed = True
                                                if def_val < 0.0001: break
                                    if def_val < 0.0001: break
                            if def_val < 0.0001: break
            if not fixed: break
        return df_pool

    pool_temp = allocated_df['Allocation Pool'].replace('', pd.NA).ffill()
    fixed_dfs = [resolve_pool(g.copy()) if (g['Remaining_Stock'] < -0.0001).any() else g for _, g in allocated_df.groupby(pool_temp)]
    df = pd.concat(fixed_dfs).sort_index()
    df['Tổng KHSX'] = df[[f"{p} - SL sau phân bổ kho" for p in product_cols if f"{p} - SL sau phân bổ kho" in df.columns]].sum(axis=1)
    return df

def export_excel(allocated_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        allocated_df.to_excel(writer, sheet_name='Allocated', index=False)
        ws = writer.sheets['Allocated']
        fill_all = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
        for col in range(1, ws.max_column + 1): ws.cell(row=1, column=col).fill = fill_all
    return output.getvalue()

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

with st.expander("2. Upload Inventory Data (5 Individual Files)"):
    col1, col2 = st.columns(2)
    with col1:
        f_kho_tot = st.file_uploader("1. Kho Tốt (Nhập xuất tồn)", type=["xlsx", "xls"])
        f_kho_clc = st.file_uploader("2. Kho CLC (Chi tiết tốt)", type=["xlsx", "xls"])
        f_tech = st.file_uploader("3. Nhà Máy Tech (TECH TỔNG)", type=["xlsx", "xls"])
    with col2:
        f_scbh = st.file_uploader("4. Nhà Máy SCBH (SCBH)", type=["xlsx", "xls"])
        f_khhv = st.file_uploader("5. KHHV (TH)", type=["xlsx", "xls"])

    if st.button("Process Inventory") and st.session_state.pivot is not None:
        with st.spinner("Processing inventory..."):
            st.session_state.pivot = process_inventory(f_kho_tot, f_kho_clc, f_tech, f_scbh, f_khhv, st.session_state.pivot)
            st.success("Inventory processed and merged successfully!")

with st.expander("3. Allocation & Download"):
    if st.session_state.pivot is not None:
        exclude = ["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity", "TONG_TON", "Tổng tồn", "Tồn kho tốt", "Tồn kho clc", "Tồn NM tech", "Tồn NM scbh", "Tồn KHHV"]
        product_cols = [c for c in st.session_state.pivot.columns if c not in exclude and 'Calculated' not in str(c)]

        cols = st.columns(len(product_cols))
        multipliers = {}
        for i, p in enumerate(product_cols):
            with cols[i]:
                multipliers[p] = st.number_input(f"Qty: {p}", min_value=0, value=1000)
                st.session_state.product_priorities[p] = st.number_input(f"Prio: {p}", min_value=1, value=i+1)

        if st.button("Run Full Allocation"):
            with st.spinner("Allocating and resolving flaws..."):
                product_cols.sort(key=lambda x: st.session_state.product_priorities.get(x, 999))
                alloc_df = run_allocation(st.session_state.pivot.copy(), product_cols, multipliers)
                alloc_df = resolve_flaws(alloc_df, product_cols)
                
                st.session_state.allocated_df = alloc_df
                st.success("Allocation complete!")

            excel_data = export_excel(st.session_state.allocated_df)
            st.download_button("Download Final Excel", data=excel_data, file_name="final_allocation_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
