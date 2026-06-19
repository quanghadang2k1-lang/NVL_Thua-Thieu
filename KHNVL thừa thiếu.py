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
            if bom_type == 'RDBOM': df['Source_RDBOM'] = uploaded_file.name; df['Level'] = df['Level'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True); df = df.ffill()
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
    merged['Level'] = merged['Level'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

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

    g_dict = {k: " | ".join([f"{r['Level']} ,{r['Source_RDBOM']}" for _, r in v.iterrows()]) for k, v in valid.groupby('Filter VNPT MAN P/N')}
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
    pivot['Level Group'] = pivot['Level Group'].astype(str)
    pivot = pivot.sort_values(by="Level Group").reset_index(drop=True)
    return processed, pivot

def process_inventory(f_tot, f_clc, f_tech, f_scbh, f_khhv, f_phu_kien_ton=None):
    dfs = []
    # 1. kho_tot
    df_tot = load_excel_header_search(f_tot, "Nhập xuất tồn", ["Mã vật tư"])
    if df_tot is not None and not df_tot.empty:
        if str(df_tot.iloc[0].values[0]).strip().startswith('('): df_tot = df_tot.iloc[1:].reset_index(drop=True)
        ma_vat_tu_cols = [col for col in df_tot.columns if 'mã vật tư' in str(col).strip().lower()]
        cuoi_ky_cols = [col for col in df_tot.columns if 'cuối kỳ' in str(col).lower() or 'tổng tồn' in str(col).lower()]
        target_cols = (ma_vat_tu_cols[:1] if ma_vat_tu_cols else []) + (cuoi_ky_cols[:1] if cuoi_ky_cols else [])
        if len(target_cols) == 2:
            df_tot = df_tot[target_cols].copy()
            df_tot.columns = ['VNPT Man P/N', 'Tồn kho tốt']
            dfs.append(df_tot)

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
    df_khhv = load_excel_header_search(f_khhv, "TH", ["Tổng"])
    if df_khhv is not None and not df_khhv.empty:
        tong_indices = [i for i, col in enumerate(df_khhv.columns) if str(col).strip().lower() == 'tổng']
        last_tong_idx = [tong_indices[-1]] if tong_indices else []
        base_indices = [i for i, col in enumerate(df_khhv.columns) if str(col).strip().lower() in ['vnpt p/n', 'vnpt pn']]
        if not base_indices:
            base_indices = [i for i, col in enumerate(df_khhv.columns) if str(col).strip().lower() in ['mã nvl', 'tên lk']]
        target_indices = base_indices[:1] + last_tong_idx
        if len(target_indices) == 2:
            df_khhv = df_khhv.iloc[:, target_indices].copy()
            df_khhv.columns = ['VNPT Man P/N', 'Tồn KHHV']
            dfs.append(df_khhv)

    # 6. phu_kien_ton
    df_phu_kien_ton = load_excel_header_search(f_phu_kien_ton, None, ["Mã VNPT 16 ký tự", "Tồn kho còn lại"])
    if df_phu_kien_ton is not None and not df_phu_kien_ton.empty:
        ma_vnpt_cols = [col for col in df_phu_kien_ton.columns if 'mã vnpt' in str(col).lower()]
        ton_kho_cols = [col for col in df_phu_kien_ton.columns if 'tồn kho' in str(col).lower()]
        target_cols = (ma_vnpt_cols[:1] if ma_vnpt_cols else []) + (ton_kho_cols[:1] if ton_kho_cols else [])
        if len(target_cols) == 2:
            df_phu_kien_ton = df_phu_kien_ton[target_cols].copy()
            df_phu_kien_ton.columns = ['VNPT Man P/N', 'Phụ kiện tồn']
            dfs.append(df_phu_kien_ton)

    if dfs:
        for i in range(len(dfs)):
            dfs[i]['VNPT Man P/N'] = dfs[i]['VNPT Man P/N'].astype(str).str.strip()
            dfs[i] = dfs[i][~dfs[i]['VNPT Man P/N'].str.lower().isin(['', 'nan', '(blank)', 'none', 'null'])]
        merged_inventory = reduce(lambda left, right: pd.merge(left, right, on='VNPT Man P/N', how='outer'), dfs)
        stock_cols = [col for col in merged_inventory.columns if col != 'VNPT Man P/N']
        for col in stock_cols:
            merged_inventory[col] = pd.to_numeric(merged_inventory[col], errors='coerce').fillna(0)
        sum_cols = [col for col in stock_cols if col not in ['Tồn kho clc', 'Phụ kiện tồn']]
        merged_inventory['Tổng tồn'] = merged_inventory[sum_cols].sum(axis=1)
        return merged_inventory
    return None

def allocate_inventory(pivot_df, product_cols):
    def get_usages(lg_str):
        if pd.isna(lg_str): return set()
        return set([part.strip() for part in str(lg_str).split('|')])

    parent = {}
    def find(i):
        if parent[i] == i: return i
        parent[i] = find(parent[i])
        return parent[i]

    def union(i, j):
        root_i = find(i)
        root_j = find(j)
        if root_i != root_j: parent[root_i] = root_j

    for idx in pivot_df.index: parent[idx] = idx

    usage_to_indices = {}
    for idx, row in pivot_df.iterrows():
        usages = get_usages(row['Level Group'])
        for u in usages:
            if u not in usage_to_indices: usage_to_indices[u] = []
            usage_to_indices[u].append(idx)

    for indices in usage_to_indices.values():
        for i in range(1, len(indices)): union(indices[0], indices[i])

    pool_dict = {}
    for idx in pivot_df.index:
        root = find(idx)
        if root not in pool_dict: pool_dict[root] = []
        pool_dict[root].append(idx)

    results = []
    pool_number = 1
    for root, indices in pool_dict.items():
        group_df = pivot_df.loc[indices].copy()
        group_df['Popularity'] = pd.to_numeric(group_df['Popularity'], errors='coerce')
        stock_col = 'Tổng tồn' if 'Tổng tồn' in group_df.columns else 'TONG_TON' if 'TONG_TON' in group_df.columns else None
        if not stock_col:
            group_df['Tổng tồn'] = 0
            stock_col = 'Tổng tồn'
        group_df[stock_col] = pd.to_numeric(group_df[stock_col], errors='coerce').fillna(0)
        group_df = group_df.sort_values(by=['Popularity', stock_col], ascending=[True, True])

        group_df['Allocation Pool'] = pool_number
        pool_number += 1

        for col in product_cols:
            if col in group_df.columns: group_df.rename(columns={col: f"{col} - Standard Qty"}, inplace=True)
            old_calc = f"{col} - Calculated"
            if old_calc in group_df.columns: group_df.rename(columns={old_calc: f"{col} - SL theo KH"}, inplace=True)

        group_remain = {}
        main_idx = group_df.index[0]

        for col in product_cols:
            calc_col = f"{col} - SL theo KH"
            group_remain[col] = group_df[calc_col].max() if calc_col in group_df.columns else 0.0
            group_df[f"{col} - SL sau phân bổ kho"] = 0.0

        for idx, row in group_df.iterrows():
            stock = row.get(stock_col, 0)
            level_group_str = str(row.get('Level Group', ''))
            for col in product_cols:
                if col in level_group_str:
                    if group_remain[col] > 0 and stock > 0:
                        use_qty = min(group_remain[col], stock)
                        group_df.at[idx, f"{col} - SL sau phân bổ kho"] = use_qty
                        group_remain[col] -= use_qty
                        stock -= use_qty
            group_df.at[idx, 'Remaining_Stock'] = stock

        for col in product_cols:
            if group_remain[col] > 0:
                valid_indices = [i for i, r in group_df.iterrows() if col in str(r.get('Level Group', ''))]
                target_idx = valid_indices[0] if valid_indices else main_idx
                group_df.at[target_idx, f"{col} - SL sau phân bổ kho"] += group_remain[col]
                group_df.at[target_idx, 'Remaining_Stock'] -= group_remain[col]
                group_remain[col] = 0.0

        results.append(group_df)

    allocated_df = pd.concat(results, ignore_index=True)

    # --- Flaw Resolution Logic ---
    def resolve_pool_flaws(df_pool):
        max_iters = 20
        for _ in range(max_iters):
            neg_mask = df_pool['Remaining_Stock'] < -0.0001
            if not neg_mask.any():
                break
            fixed_something = False
            for deficit_idx in df_pool[neg_mask].index:
                deficit_val = -df_pool.loc[deficit_idx, 'Remaining_Stock']
                for p_col in product_cols:
                    alloc_col = f"{p_col} - SL sau phân bổ kho"
                    if alloc_col in df_pool.columns and df_pool.loc[deficit_idx, alloc_col] > 0:
                        valid_for_A = [idx for idx, row in df_pool.iterrows() if p_col in str(row.get('Level Group', ''))]
                        for shared_idx in valid_for_A:
                            if shared_idx == deficit_idx: continue
                            for p_col_B in product_cols:
                                if p_col_B == p_col: continue
                                alloc_col_B = f"{p_col_B} - SL sau phân bổ kho"
                                if alloc_col_B in df_pool.columns:
                                    b_taken = df_pool.loc[shared_idx, alloc_col_B]
                                    if b_taken > 0:
                                        valid_for_B = [idx for idx, row in df_pool.iterrows() if p_col_B in str(row.get('Level Group', ''))]
                                        for spare_idx in valid_for_B:
                                            if spare_idx == shared_idx or spare_idx == deficit_idx: continue
                                            spare_stock = df_pool.loc[spare_idx, 'Remaining_Stock']
                                            if spare_stock > 0:
                                                shift_amount = min(deficit_val, b_taken, spare_stock)
                                                if shift_amount > 0:
                                                    df_pool.loc[shared_idx, alloc_col_B] -= shift_amount
                                                    df_pool.loc[spare_idx, alloc_col_B] += shift_amount
                                                    df_pool.loc[spare_idx, 'Remaining_Stock'] -= shift_amount

                                                    df_pool.loc[deficit_idx, alloc_col] -= shift_amount
                                                    df_pool.loc[shared_idx, alloc_col] += shift_amount
                                                    df_pool.loc[deficit_idx, 'Remaining_Stock'] += shift_amount

                                                    deficit_val -= shift_amount
                                                    fixed_something = True
                                                    if deficit_val < 0.0001: break
                                        if deficit_val < 0.0001: break
                            if deficit_val < 0.0001: break
            if not fixed_something:
                break

        neg_mask_final = df_pool['Remaining_Stock'] < -0.0001
        if neg_mask_final.any():
            for deficit_idx in df_pool[neg_mask_final].index:
                for p_col in product_cols:
                    alloc_col = f"{p_col} - SL sau phân bổ kho"
                    if alloc_col in df_pool.columns and df_pool.loc[deficit_idx, alloc_col] > 0:
                        deficit_val = -df_pool.loc[deficit_idx, 'Remaining_Stock']
                        if deficit_val <= 0.0001: continue
                        valid_for_A = [idx for idx, row in df_pool.iterrows() if p_col in str(row.get('Level Group', ''))]
                        if valid_for_A:
                            highest_pop_idx = max(valid_for_A, key=lambda i: pd.to_numeric(df_pool.loc[i, 'Popularity'], errors='coerce') if pd.notna(df_pool.loc[i, 'Popularity']) else 0)
                            if highest_pop_idx != deficit_idx:
                                shift_amount = min(deficit_val, df_pool.loc[deficit_idx, alloc_col])
                                if shift_amount > 0:
                                    df_pool.loc[deficit_idx, alloc_col] -= shift_amount
                                    df_pool.loc[highest_pop_idx, alloc_col] += shift_amount
                                    df_pool.loc[deficit_idx, 'Remaining_Stock'] += shift_amount
                                    df_pool.loc[highest_pop_idx, 'Remaining_Stock'] -= shift_amount
        return df_pool

    pool_temp = allocated_df['Allocation Pool'].replace('', pd.NA).ffill()
    fixed_dfs = []
    for pool_id, group in allocated_df.groupby(pool_temp):
        if (group['Remaining_Stock'] < -0.0001).any():
            fixed_group = resolve_pool_flaws(group.copy())
            fixed_dfs.append(fixed_group)
        else:
            fixed_dfs.append(group)
    allocated_df = pd.concat(fixed_dfs).sort_index()
    # --- End of Flaw Resolution Logic ---

    all_cols = list(allocated_df.columns)
    fixed_start_cols = [c for c in ['Level Group', 'Allocation Pool', 'Filter VNPT MAN P/N', 'Description', 'Popularity'] if c in all_cols]

    prod_cols_ordered = []
    for p_col in product_cols:
        for suffix in [" - Standard Qty", " - SL theo KH", " - SL sau phân bổ kho"]:
            if f"{p_col}{suffix}" in all_cols: prod_cols_ordered.append(f"{p_col}{suffix}")

    allocated_df['Tổng KHSX'] = allocated_df[[c for c in all_cols if 'SL sau phân bổ kho' in c]].sum(axis=1)
    if 'Tồn kho clc' in allocated_df.columns:
        allocated_df['Tồn Active'] = allocated_df['Remaining_Stock'] - pd.to_numeric(allocated_df['Tồn kho clc'], errors='coerce').fillna(0)
    if 'Tồn kho clc' in allocated_df.columns:
        allocated_df['Tồn Active'] = allocated_df['Remaining_Stock'] - pd.to_numeric(allocated_df['Phụ kiện tồn'], errors='coerce').fillna(0)

    end_cols = [c for c in ['Tổng KHSX', 'Tồn kho tốt', 'Tồn kho clc', 'Tồn NM tech', 'Tồn NM scbh', 'Tồn KHHV', 'Phụ kiện tồn', 'Tổng tồn', 'Remaining_Stock', 'Tồn Active'] if c in allocated_df.columns]

    final_ordered_cols = fixed_start_cols + prod_cols_ordered + end_cols
    final_ordered_cols = [c for c in final_ordered_cols if c in allocated_df.columns]
    allocated_df = allocated_df[final_ordered_cols]

    allocated_df['Allocation Pool'] = allocated_df['Allocation Pool'].astype(str)
    allocated_df.loc[allocated_df['Allocation Pool'].duplicated(), 'Allocation Pool'] = ''
    return allocated_df

def generate_excel(allocated_df, processed_df=None, pivot=None, merged_inventory=None, summary_df=None):
    # Reorder columns to group by metric instead of product
    all_cols = list(allocated_df.columns)
    fixed_start = [c for c in ['Level Group', 'Allocation Pool', 'Filter VNPT MAN P/N', 'Description', 'Popularity'] if c in all_cols]
    std_cols_name = [c for c in all_cols if ' - Standard Qty' in c]
    kh_cols_name = [c for c in all_cols if ' - SL theo KH' in c]
    alloc_cols_name = [c for c in all_cols if ' - SL sau phân bổ kho' in c]

    # Preserve any remaining columns at the end
    grouped_set = set(fixed_start + std_cols_name + kh_cols_name + alloc_cols_name)
    end_cols = [c for c in all_cols if c not in grouped_set]

    new_order = fixed_start + std_cols_name + kh_cols_name + alloc_cols_name + end_cols
    allocated_df = allocated_df[new_order]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        allocated_df.to_excel(writer, sheet_name='Allocated', index=False)

        if processed_df is not None:
            processed_df.to_excel(writer, sheet_name='BOM Result', index=False)
        if pivot is not None:
            pivot.to_excel(writer, sheet_name='Pivot', index=False)
        if merged_inventory is not None:
            merged_inventory.to_excel(writer, sheet_name='Inventory', index=False)
        if summary_df is not None:
            summary_df.to_excel(writer, sheet_name='KHSX', index=False)

        worksheet = writer.sheets['Allocated']
        alloc_pool_col_idx = None
        max_col = worksheet.max_column

        std_cols = []
        kh_cols = []
        alloc_cols = []

        for col_idx, col_name in enumerate(allocated_df.columns, 1):
            col_str = str(col_name)
            if col_str == 'Allocation Pool':
                alloc_pool_col_idx = col_idx
            if ' - Standard Qty' in col_str:
                std_cols.append(col_idx)
            elif ' - SL theo KH' in col_str:
                kh_cols.append(col_idx)
            elif ' - SL sau phân bổ kho' in col_str:
                alloc_cols.append(col_idx)

        fill_all = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
        fill_std = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
        fill_kh = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        fill_alloc = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        for col_idx in range(1, max_col + 1):
            worksheet.cell(row=1, column=col_idx).fill = fill_all
        for col_idx in std_cols:
            worksheet.cell(row=1, column=col_idx).fill = fill_std
        for col_idx in kh_cols:
            worksheet.cell(row=1, column=col_idx).fill = fill_kh
        for col_idx in alloc_cols:
            worksheet.cell(row=1, column=col_idx).fill = fill_alloc

        sep_cols = []
        if std_cols:
            sep_cols.append(min(std_cols) - 1)
            sep_cols.append(max(std_cols))
        if kh_cols:
            sep_cols.append(max(kh_cols))
        if alloc_cols:
            sep_cols.append(max(alloc_cols))

        medium_side = Side(style='medium')

        for r in range(1, len(allocated_df) + 2):
            for c in sep_cols:
                if c > 0:
                    cell = worksheet.cell(row=r, column=c)
                    cell.border = Border(left=cell.border.left, right=medium_side, top=cell.border.top, bottom=cell.border.bottom)

        if alloc_pool_col_idx:
            center_alignment = Alignment(horizontal='center', vertical='center')
            medium_bottom_border = Side(style='medium')

            for r in range(1, len(allocated_df) + 2):
                worksheet.cell(row=r, column=alloc_pool_col_idx).alignment = center_alignment

            start_row = 2
            for i in range(1, len(allocated_df)):
                val = allocated_df['Allocation Pool'].iloc[i]
                current_excel_row = i + 2
                if val != '':
                    prev_row = current_excel_row - 1
                    if prev_row > start_row:
                        worksheet.merge_cells(start_row=start_row, start_column=alloc_pool_col_idx, end_row=prev_row, end_column=alloc_pool_col_idx)
                    for col in range(1, max_col + 1):
                        cell = worksheet.cell(row=prev_row, column=col)
                        cell.border = Border(left=cell.border.left, right=cell.border.right, top=cell.border.top, bottom=medium_bottom_border)
                    start_row = current_excel_row

            last_row = len(allocated_df) + 1
            if last_row > start_row:
                worksheet.merge_cells(start_row=start_row, start_column=alloc_pool_col_idx, end_row=last_row, end_column=alloc_pool_col_idx)
            for col in range(1, max_col + 1):
                cell = worksheet.cell(row=last_row, column=col)
                cell.border = Border(left=cell.border.left, right=cell.border.right, top=cell.border.top, bottom=medium_bottom_border)

    output.seek(0)
    return output

# --- UI Workflow ---
with st.expander("1. Upload BOM Files", expanded=True):
    rdbom_files = st.file_uploader("Upload RDBOM.xlsx (Required)", type=["xlsx", "xls"], accept_multiple_files=True)
    manbom_files = st.file_uploader("Upload MANBOM.xlsx (Optional)", type=["xlsx", "xls"], accept_multiple_files=True)
    if st.button("Xử lí BOMs"):
        with st.spinner("Đang xử lí..."):
            processed, pivot = process_boms(rdbom_files, manbom_files)
            if pivot is not None:
                st.session_state.processed_df, st.session_state.pivot = processed, pivot
                st.success("BOMs đã xử lí thành công!")
                st.markdown("### Bảng kết quả BOM")
                st.dataframe(st.session_state.processed_df)
                st.markdown("### Bảng Pivot")
                st.dataframe(st.session_state.pivot)

with st.expander("2. Kế hoạch sản xuất & Nhu cầu sản lượng", expanded=True):
    if st.session_state.pivot is not None:
        pivot_df = st.session_state.pivot.copy()
        index_cols = ["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity"]
        rdbom_cols = [col for col in pivot_df.columns if col not in index_cols]
        st.session_state.product_cols = rdbom_cols
        cols = st.columns(3)
        multipliers = {}
        for i, col in enumerate(rdbom_cols):
            with cols[i % 3]:
                multipliers[col] = st.number_input(f"Multiplier for {col}", min_value=0.0, value=1.0, step=1.0)
        if st.button("Tính nhu cầu"):
            for col in rdbom_cols:
                pivot_df[f"{col} - Calculated"] = pivot_df[col] * multipliers[col]
            st.session_state.pivot_calculated = pivot_df
            st.success("Đã tính xong nhu cầu!")
            st.markdown("### Bảng pivot kèm sản lượng theo KHSX")
            st.dataframe(st.session_state.pivot_calculated)
    else:
        st.info("Hãy hoàn thành bước 1.")

with st.expander("3. Upload Files Tồn bộ phận", expanded=True):
    if st.session_state.pivot_calculated is not None:
        col1, col2 = st.columns(2)
        with col1:
            f_tot = st.file_uploader("Upload 'Kho Tốt'", type=["xlsx", "xls"])
            f_clc = st.file_uploader("Upload 'Kho CLC'", type=["xlsx", "xls"])
            f_tech = st.file_uploader("Upload 'Nhà máy Tech'", type=["xlsx", "xls"])
        with col2:
            f_scbh = st.file_uploader("Upload 'Nhà máy SCBH'", type=["xlsx", "xls"])
            f_khhv = st.file_uploader("Upload 'KHHV'", type=["xlsx", "xls"])
            f_phu_kien_ton = st.file_uploader("Upload 'Phụ kiện tồn'", type=["xlsx", "xls"])
        if st.button("Xử lí tồn BP"):
            with st.spinner("Đang xử lí..."):
                merged_inv = process_inventory(f_tot, f_clc, f_tech, f_scbh, f_khhv, f_phu_kien_ton)
                if merged_inv is not None:
                    st.session_state.merged_inventory = merged_inv
                    pivot_final = st.session_state.pivot_calculated.copy()
                    cols_to_merge = [c for c in ['VNPT Man P/N', 'Tồn kho tốt', 'Tồn kho clc', 'Tồn NM tech', 'Tồn NM scbh', 'Tồn KHHV', 'Phụ kiện tồn', 'Tổng tồn'] if c in merged_inv.columns]
                    pivot_final = pd.merge(pivot_final, merged_inv[cols_to_merge], left_on='Filter VNPT MAN P/N', right_on='VNPT Man P/N', how='left')
                    if 'VNPT Man P/N' in pivot_final.columns: pivot_final = pivot_final.drop(columns=['VNPT Man P/N'])
                    st.session_state.pivot_calculated = pivot_final
                    st.success("Kết hợp dữ liệu tồn BP thành công!")
                    st.markdown("### Dữ liệu tồn")
                    st.dataframe(st.session_state.merged_inventory)
                    st.markdown("### Bảng Pivot kèm tồn BP ")
                    st.dataframe(st.session_state.pivot_calculated)
                else:
                    st.warning("Không thể trích xuất tồn BP.")
    else:
        st.info("Hãy tính nhu cầu sản lượng ở bước 2.")

with st.expander("4. Mức độ ưu tiên sản phẩm & Phân bổ linh kiện ", expanded=True):
    if st.session_state.pivot_calculated is not None and 'Tổng tồn' in st.session_state.pivot_calculated.columns:
        st.markdown("### Đặt mức độ ưu tiên cho sản phẩm (Số nhỏ = Mức độ ưu tiên cao)")
        st.markdown("#### Linh kiện tồn sẽ được phân bổ theo thứ tự ưu tiên của sản phẩm")
        st.markdown("##### Để nguyên nếu không muốn đặt mức độ ưu tiên")
        cols = st.columns(3)
        priorities = {}
        for i, col in enumerate(st.session_state.product_cols):
            with cols[i % 3]:
                priorities[col] = st.number_input('Priority of '+ {col} , min_value=1, value = i + 1, step=1)

        if st.button("Phân bổ"):
            with st.spinner("Đang phân bổ linh kiện..."):
                st.session_state.product_priorities = priorities
                sorted_products = sorted(st.session_state.product_cols, key=lambda x: priorities.get(x, 999))

                allocated = allocate_inventory(st.session_state.pivot_calculated, sorted_products)
                st.session_state.allocated_df = allocated
                st.success("Phân bổ linh kiện thành công!")
                st.markdown("### Dữ liệu phân bổ cuối cùng")
                st.dataframe(st.session_state.allocated_df)
    else:
        st.info("Hãy upload dữ liều tồn bộ phận và đảm bảo cột 'Tổng tồn' đã được tính ở bước 3.")

#Download final result
if st.session_state.get('allocated_df') is not None:
    st.markdown("### 5. Download Output")

    # Generate summary dataframe
    summary_data = []
    p_cols = st.session_state.product_cols
    for p in p_cols:
        priority = st.session_state.product_priorities.get(p, "")
        multiplier = 0
        std_col = f"{p} - Standard Qty"
        calc_col = f"{p} - SL theo KH"

        if std_col in st.session_state.allocated_df.columns and calc_col in st.session_state.allocated_df.columns:
            valid_rows = st.session_state.allocated_df[st.session_state.allocated_df[std_col] > 0]
            if not valid_rows.empty:
                multiplier = valid_rows.iloc[0][calc_col] / valid_rows.iloc[0][std_col]
                multiplier = round(multiplier, 4)
                if multiplier.is_integer():
                    multiplier = int(multiplier)

        summary_data.append({
            'Tên sản phẩm': p,
            'Sản lượng (KHSX)': multiplier,
            'Mức độ ưu tiên': priority
        })

    summary_df = pd.DataFrame(summary_data)

    with st.spinner("Preparing Excel file for download..."):
        excel_bytes = generate_excel(
            st.session_state.allocated_df,
            st.session_state.get('processed_df'),
            st.session_state.get('pivot'),
            st.session_state.get('merged_inventory'),
            summary_df
        )

    st.download_button(
        label="📥 Tải file kết quả cuối cùng",
        data=excel_bytes,
        file_name="final_allocation_result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
