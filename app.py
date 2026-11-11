import streamlit as st
import pandas as pd
import re

# ---------------------------------------------------------------------------
# Page Configuration & Custom CSS
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Grade Checker Tool", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .sub-text {
        color: #4B5563;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    .stAlert {
        border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# App Interface
# ---------------------------------------------------------------------------
st.markdown('<p class="main-header">📊 Blackboard vs Internal Grade Checker</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-text">Upload your Blackboard export and internal grade sheets below to automatically detect mismatches.</p>',
    unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Blackboard Export")
    bb_file = st.file_uploader("Upload Blackboard File (.csv, .xlsx, .xls, .txt)", type=["csv", "xlsx", "xls", "txt"],
                               key="bb")

with col2:
    st.subheader("2. Internal Grade Sheet")
    internal_file = st.file_uploader("Upload Internal Sheet (.csv, .xlsx, .xls)", type=["csv", "xlsx", "xls"],
                                     key="internal")

if bb_file and internal_file:
    # Universal file reader function
    def load_file(uploaded_file):
        filename = uploaded_file.name.lower()
        if filename.endswith('.csv'):
            try:
                return pd.read_csv(uploaded_file)
            except:
                return pd.read_csv(uploaded_file, sep="\t", encoding="utf-16")
        elif filename.endswith(('.xlsx', '.xls')):
            try:
                return pd.read_excel(uploaded_file, engine="openpyxl")
            except:
                return pd.read_excel(uploaded_file, engine="xlrd")
        else:
            return pd.read_csv(uploaded_file, sep="\t")


    try:
        df_bb = load_file(bb_file)
        df_internal = load_file(internal_file)

        # Clean headers
        df_bb.columns = df_bb.columns.astype(str).str.strip()
        df_internal.columns = df_internal.columns.astype(str).str.strip()


        # Identify ID columns
        def find_id_col(df):
            for target in ['student id', 'cand key', '#cand key', 'username', 'candidate id']:
                for col in df.columns:
                    if target in col.lower():
                        return col
            return df.columns[0]


        bb_id_col = find_id_col(df_bb)
        internal_id_col = find_id_col(df_internal)

        # Pick assessment column
        bb_score_cols = [c for c in df_bb.columns if
                         'Score' in c or 'Pts' in c or 'Reassessment' in c or 'Assessment' in c]

        st.divider()
        st.subheader("3. Configuration")
        chosen_bb_col = st.selectbox("Select Blackboard Assessment Column:",
                                     bb_score_cols if bb_score_cols else df_bb.columns)

        internal_mark_col = next((c for c in df_internal.columns if 'mark' in c.lower() or 'grade' in c.lower()),
                                 df_internal.columns[1] if len(df_internal.columns) > 1 else df_internal.columns[0])
        st.info(f"Targeting Internal Sheet Mark Column: **{internal_mark_col}**")

        if st.button("Run Grade Comparison", type="primary"):
            # Standardize IDs
            def clean_id(val):
                if pd.isna(val): return ""
                s = str(val).strip()
                digits = re.sub(r'\D', '', s)
                return digits.lstrip('0') if digits else s.lower()


            df_bb['Clean_ID'] = df_bb[bb_id_col].apply(clean_id)
            df_internal['Clean_ID'] = df_internal[internal_id_col].apply(clean_id)

            # Subsets & Merge
            bb_subset = df_bb[['Clean_ID', bb_id_col, chosen_bb_col]].rename(
                columns={bb_id_col: 'BB_ID', chosen_bb_col: 'BB_Score'})
            internal_subset = df_internal[['Clean_ID', internal_id_col, internal_mark_col]].rename(
                columns={internal_id_col: 'Internal_ID', internal_mark_col: 'Internal_Mark'})

            merged = pd.merge(bb_subset, internal_subset, on='Clean_ID', how='inner')

            merged['BB_Num'] = pd.to_numeric(merged['BB_Score'], errors='coerce')
            merged['Int_Num'] = pd.to_numeric(merged['Internal_Mark'], errors='coerce')

            mismatches = merged[
                (merged['BB_Num'] != merged['Int_Num']) &
                ~(merged['BB_Num'].isna() & merged['Int_Num'].isna())
                ]

            st.divider()
            st.subheader("📊 Results Summary")

            col_res1, col_res2 = st.columns(2)
            col_res1.metric("Total Students Matched", len(merged))
            col_res2.metric("Total Grade Mismatches Found", len(mismatches), delta_color="inverse")

            if not mismatches.empty:
                st.warning(f"Found {len(mismatches)} grade discrepancies!")
                st.dataframe(mismatches[['BB_ID', 'BB_Score', 'Internal_Mark']], use_container_width=True)

                # Download Button for Excel Report
                output_file = "grade_mismatches_report.xlsx"
                mismatches.drop(columns=['Clean_ID', 'BB_Num', 'Int_Num'], errors='ignore').to_excel(output_file,
                                                                                                     index=False)

                with open(output_file, "rb") as f:
                    st.download_button("📥 Download Mismatch Report (Excel)", f, file_name=output_file,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.chssheet")
            else:
                st.success("🎉 Perfect! No grade mismatches found between sheets.")

    except Exception as e:
        st.error(f"An error occurred while processing the files: {e}")