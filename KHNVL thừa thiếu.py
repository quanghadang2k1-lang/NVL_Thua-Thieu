import streamlit as st
import pandas as pd
import io
import re

# --- Streamlit App Configuration ---
st.set_page_config(layout="wide", page_title="BOM Allocation Tool")
st.title("BOM Matching & Stock Distribution Tool")
st.markdown("Upload your BOMs, Inventory, and define your production plan to allocate stock.")

# --- Session State Initialization ---
def init_state(key, default):
    if key not in st.session_state: st.session_state[key] = default

init_state('rdbom_df', None)
init_state('manbom_df', None)
init_state('merged_df', None)
init_state('processed_df', None)
init_state('pivot', None)
init_state('merged_inventory', None)
init_state('product_cols', [])
init_state('allocated_df', None)
init_state('level_groups_dict', {})
init_state('product_priorities', {})

# --- Helper Functions ---
def load_excel_header_search(uploaded_file, sheet_keyword, keywords, is_bom=False, bom_type='RDBOM'):
    if uploaded_file is None: return None
    content_io = io.BytesIO(uploaded_file.getvalue())
    file_extension = uploaded_file.name.lower().split('.')[-1]
    engine_to_use = 'openpyxl' # Default to openpyxl

    if file_extension == 'xls':
        try:
            # Try to import xlrd. If successful, use it.
            import xlrd # Attempt import
            if xlrd.__version__ >= '2.0.1': # Check version for .xls support
                engine_to_use = 'xlrd'
            else:
                st.warning(f"xlrd version {xlrd.__version__} is installed, but version >= 2.0.1 is recommended for .xls files. Using openpyxl as a fallback for '{uploaded_file.name}'.")
        except ImportError:
            st.warning(f"xlrd library not found or failed to import. Attempting to open '{uploaded_file.name}' with openpyxl (may have limited support for older .xls formats).")
        except Exception as e:
            st.warning(f"An unexpected error occurred while checking xlrd for '{uploaded_file.name}': {e}. Attempting to open with openpyxl.")

    kwargs = {'engine': engine_to_use}
    if engine_to_use == 'xlrd':
        kwargs['engine_kwargs'] = {'ignore_workbook_corruption': True}

    try:
        xls = pd.ExcelFile(content_io, **kwargs)
        sheet_name = next((sn for sn in xls.sheet_names if str(sheet_keyword).lower() in str(sn).lower()), None) if sheet_keyword else xls.sheet_names[0]
        if sheet_name is None:
            st.error(f"No sheet matching '{sheet_keyword}' found in '{uploaded_file.name}'.")
            return None

        content_io.seek(0) # Reset stream for read_excel
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
        st.error(f"Error loading {uploaded_file.name} with engine '{engine_to_use}': {str(e)}")
        return None

# --- Step 1: BOM Upload & Merge ---
with st.expander("1. Upload BOM Files & Process", expanded=True):
    rdbom_files = st.file_uploader("Upload RDBOM.xlsx (Required)", type=["xlsx", "xls"], accept_multiple_files=True)
    manbom_files = st.file_uploader("Upload MANBOM.xlsx (Optional)", type=["xlsx", "xls"], accept_multiple_files=True)

    if st.button("Process BOMs"):
        with st.spinner("Processing..."):
            if rdbom_files:
                st.session_state.rdbom_df = pd.concat([load_excel_header_search(f, None, ['Level', 'VNPT P/N'], True, 'RDBOM') for f in rdbom_files], ignore_index=True)
            if manbom_files:
                st.session_state.manbom_df = pd.concat([load_excel_header_search(f, None, ['VNPT P/N', 'Tỉ lệ tiêu hao'], True, 'MANBOM') for f in manbom_files], ignore_index=True)

            if st.session_state.rdbom_df is not None:
                rdbom = st.session_state.rdbom_df.copy()
                rdbom['Base_Project'] = rdbom['Source_RDBOM'].str.replace(r'\s*\(\d+\)', '', regex=True).str.replace(r'\.xls[x]?$', '', regex=True, flags=re.IGNORECASE)
                if st.session_state.manbom_df is not None:
                    manbom = st.session_state.manbom_df.copy()
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
                merged = merged[[c for c in ['Source_RDBOM', 'Source_MANBOM', 'VNPT P/N', 'Level', 'Description', 'VNPT MAN P/N', 'Quantity/Product', 'consumption rate', 'Standard quantity'] if c in merged.columns]]

                # Profiling
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
                st.session_state.processed_df = processed

                # Ensure required pivot columns exist
                if 'Description' not in processed.columns:
                    processed['Description'] = ''
                if 'Source_RDBOM' not in processed.columns:
                    processed['Source_RDBOM'] = 'Unknown'
                if 'Standard quantity' not in processed.columns:
                    processed['Standard quantity'] = 0

                # Pivot
                pivot = pd.pivot_table(processed, index=["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity"], columns=["Source_RDBOM"], values="Standard quantity", aggfunc="sum", fill_value=0).reset_index()
                st.session_state.pivot = pivot.sort_values(by="Level Group").reset_index(drop=True)
                st.success("BOMs processed successfully!")

# --- Step 2: Inventory Upload ---
with st.expander("2. Upload Inventory Data"):
    stock_file = st.file_uploader("Upload Consolidated Inventory File (Stocks .xlsx)", type=["xlsx", "xls"])
    if st.button("Process Inventory") and stock_file and st.session_state.pivot is not None:
        stocks = pd.read_excel(stock_file)
        if 'Level Group' in stocks.columns: stocks = stocks.drop(columns=['Level Group'])
        if 'Filter VNPT MAN P/N' in stocks.columns:
            st.session_state.pivot = pd.merge(st.session_state.pivot, stocks, on='Filter VNPT MAN P/N', how='left')
            st.success("Inventory merged!")

# --- Step 3: Allocation ---
with st.expander("3. Allocation & Download"):
    if st.session_state.pivot is not None:
        exclude = ["Level Group", "Filter VNPT MAN P/N", "Description", "Popularity", "TONG_TON", "Tổng tồn", "Tồn kho tốt", "Tồn kho clc", "Tồn NM tech", "Tồn NM scbh", "Tồn KHHV"]
        product_cols = [c for c in st.session_state.pivot.columns if c not in exclude and 'Calculated' not in str(c)]

        st.write("Set Production Multipliers and Priorities (Lower number = Higher Priority):")
        cols = st.columns(len(product_cols))
        multipliers = {}
        for i, p in enumerate(product_cols):
            with cols[i]:
                multipliers[p] = st.number_input(f"Qty: {p}", min_value=0, value=1000)
                st.session_state.product_priorities[p] = st.number_input(f"Prio: {p}", min_value=1, value=i+1)

        if st.button("Run Full Allocation"):
            pivot = st.session_state.pivot.copy()
            product_cols.sort(key=lambda x: st.session_state.product_priorities.get(x, 999))
            st.session_state.product_cols = product_cols

            for p in product_cols: pivot[f"{p} - Calculated"] = pivot[p] * multipliers[p]

            # Union-Find Grouping
            parent = {i: i for i in pivot.index}
            def find(i):
                if parent[i] == i: return i
                parent[i] = find(parent[i]); return parent[i]
            def union(i, j):
                root_i, root_j = find(i), find(j)
                if root_i != root_j: parent[root_i] = root_j

            usage_to_indices = {}
            for idx, row in pivot.iterrows():
                usages = set([x.strip() for x in str(row['Level Group']).split('|')]) if pd.notna(row['Level Group']) else set()
                for u in usages:
                    usage_to_indices.setdefault(u, []).append(idx)
            for indices in usage_to_indices.values():
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
                    lg = str(row.get('Level Group', ''))
                    for p in product_cols:
                        if p in lg and group_remain[p] > 0 and stock > 0:
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

            alloc_df = pd.concat(results, ignore_index=True)
            alloc_df['Tổng KHSX'] = alloc_df[[f"{p} - SL sau phân bổ kho" for p in product_cols]].sum(axis=1)
            st.session_state.allocated_df = alloc_df
            st.success("Allocation complete!")

            # Export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                alloc_df.to_excel(writer, sheet_name='Allocated', index=False)
            st.download_button("Download Final Excel", data=output.getvalue(), file_name="final_allocation_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
