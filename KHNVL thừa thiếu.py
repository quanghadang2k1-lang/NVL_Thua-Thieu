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
            if bom_type == 'RDBOM': 
                df['Source_RDBOM'] = uploaded_file.name
                df['Level'] = df['Level'].astype(str).str.replace(r'\.0$', '', regex=True)
                df = df.ffill()
            if bom_type == 'MANBOM': 
                df['Source_MANBOM'] = uploaded_file.name
        return df
    except Exception as e:
        st.error(f"Error loading {uploaded_file.name}: {str(e)}")
        return None

def process_boms(rdbom_files, manbom_files):
    rdbom_df = pd.concat([load_excel_header_search(f, None, ['level', 'vnpt p/n'], True, 'RDBOM') for f in rdbom_files], ignore_index=True) if rdbom_files else None
    manbom_df = pd.concat([load_excel_header_search(f, None, ['vnpt p/n', 'tỉ lệ tiêu hao'], True, 'MANBOM') for f in manbom_files], ignore_index=True) if manbom_files else None

    if rdbom_df is None: return None

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
    merged['Level'] = merged['Level'].astype(str).str.replace(r'\.0$', '', regex=True)

    if 'Tỉ lệ tiêu hao' in merged.columns: merged = merged.rename(columns={'Tỉ lệ tiêu hao': 'consumption rate'})
    for col in ['Quantity / Product ', 'Quantity / Product']:
        if col in merged.columns: merged = merged.rename(columns={col: 'Quantity/Product'})
    merged['consumption rate'] = pd.to_numeric(merged.get('consumption rate', 0), errors='coerce').fillna(0)
    merged['Quantity/Product'] = pd.to_numeric(merged.get('Quantity/Product', 0), errors='coerce').fillna(0)
    merged['Standard quantity'] = merged['Quantity/Product'] + merged['consumption rate']

    return merged

# --- UI Workflow ---
with st.expander("1. Upload BOM Files & Process", expanded=True):
    rdbom_files = st.file_uploader("Upload RDBOM.xlsx (Required)", type=["xlsx", "xls"], accept_multiple_files=True)
    manbom_files = st.file_uploader("Upload MANBOM.xlsx (Optional)", type=["xlsx", "xls"], accept_multiple_files=True)
    
    if st.button("Process BOMs"):
        if not rdbom_files:
            st.warning("Please upload at least one RDBOM file.")
        else:
            with st.spinner("Processing & Merging..."):
                merged_df = process_boms(rdbom_files, manbom_files)
                if merged_df is not None:
                    st.session_state.processed_df = merged_df
                    st.success("BOMs processed and merged successfully!")
                    st.markdown("### Merged BOM Results")
                    st.dataframe(st.session_state.processed_df)
