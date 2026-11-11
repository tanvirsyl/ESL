import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ---------------------------------------------------------------------------
# 1. Open File Dialogs
# ---------------------------------------------------------------------------
root = tk.Tk()
root.withdraw()

print("Please select your Blackboard export file...")
blackboard_file = filedialog.askopenfilename(
    title="Select Blackboard Export File",
    filetypes=[("All Files", "*.*"), ("Excel & CSV", "*.xlsx *.xls *.csv *.txt")]
)

if not blackboard_file:
    print("No Blackboard file selected. Exiting.")
    exit()

print("Please select your Internal Grade Sheet file...")
internal_file = filedialog.askopenfilename(
    title="Select Internal Grade Sheet File",
    filetypes=[("All Files", "*.*"), ("Excel & CSV", "*.xlsx *.xls *.csv *.txt")]
)

if not internal_file:
    print("No Internal Grade Sheet file selected. Exiting.")
    exit()


# ---------------------------------------------------------------------------
# 2. Universal File Reader Function
# ---------------------------------------------------------------------------
def read_sheet(file_path):
    try:
        return pd.read_csv(file_path)
    except Exception:
        pass
    try:
        return pd.read_csv(file_path, sep="\t", encoding="utf-16")
    except Exception:
        pass
    try:
        return pd.read_csv(file_path, sep="\t")
    except Exception:
        pass
    try:
        return pd.read_excel(file_path, engine="openpyxl")
    except Exception:
        pass
    try:
        return pd.read_excel(file_path, engine="xlrd")
    except Exception:
        pass

    tables = pd.read_html(file_path)
    return tables[0]


df_bb = read_sheet(blackboard_file)
df_internal = read_sheet(internal_file)

# Clean column headers
df_bb.columns = df_bb.columns.astype(str).str.strip()
df_internal.columns = df_internal.columns.astype(str).str.strip()


# ---------------------------------------------------------------------------
# 3. Dynamic Column Matching Helper
# ---------------------------------------------------------------------------
def find_id_column(df, file_label):
    targets = ['student id', 'cand key', '#cand key', 'username', 'candidate id']
    for target in targets:
        for col in df.columns:
            if target in col.lower():
                return col
    print(f"\nAvailable columns in {file_label}: {list(df.columns)}")
    raise KeyError(f"Could not automatically detect Student ID/Candidate Key column in {file_label}.")


bb_id_col = find_id_column(df_bb, "Blackboard Export")
internal_id_col = find_id_column(df_internal, "Internal Grade Sheet")

print(f"\n[OK] Blackboard ID Column: '{bb_id_col}'")
print(f"[OK] Internal Sheet ID Column: '{internal_id_col}'")

# ---------------------------------------------------------------------------
# 4. Choose Blackboard Score Column
# ---------------------------------------------------------------------------
bb_score_cols = [c for c in df_bb.columns if 'Score' in c or 'Pts' in c or 'Reassessment' in c or 'Assessment' in c]

if not bb_score_cols:
    raise ValueError("No score/assessment columns found in Blackboard file.")

print("\nAvailable Score Columns found in Blackboard:")
for i, col in enumerate(bb_score_cols):
    print(f"  [{i}] {col}")

choice = input(f"\nEnter column index to compare [0-{len(bb_score_cols) - 1}] (default 0): ").strip()
try:
    chosen_idx = int(choice)
    chosen_bb_col = bb_score_cols[chosen_idx]
except Exception:
    chosen_bb_col = bb_score_cols[0]

# Detect internal mark column
internal_mark_col = None
for col in df_internal.columns:
    if col.lower() == 'mark' or 'mark' in col.lower() or 'grade' in col.lower():
        internal_mark_col = col
        break

if not internal_mark_col:
    print(f"\nAvailable columns in Internal Sheet: {list(df_internal.columns)}")
    raise KeyError("Could not find a 'Mark' or 'Grade' column in the internal sheet.")

print(f"\n[OK] Blackboard Score Column: '{chosen_bb_col}'")
print(f"[OK] Internal Mark Column: '{internal_mark_col}'")


# ---------------------------------------------------------------------------
# 5. Standardize Student IDs
# ---------------------------------------------------------------------------
def clean_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    digits = re.sub(r'\D', '', s)
    if digits:
        return digits.lstrip('0')
    return s.lower()


df_bb['Clean_ID'] = df_bb[bb_id_col].apply(clean_id)
df_internal['Clean_ID'] = df_internal[internal_id_col].apply(clean_id)

# ---------------------------------------------------------------------------
# 6. Merge and Find Mismatches
# ---------------------------------------------------------------------------
bb_cols_to_keep = ['Clean_ID', bb_id_col, chosen_bb_col]
if 'Last Name' in df_bb.columns: bb_cols_to_keep.append('Last Name')
if 'First Name' in df_bb.columns: bb_cols_to_keep.append('First Name')

bb_subset = df_bb[bb_cols_to_keep].copy()
bb_subset = bb_subset.rename(columns={bb_id_col: 'BB_Student_ID', chosen_bb_col: 'BB_Score'})

internal_subset = df_internal[['Clean_ID', internal_id_col, internal_mark_col]].copy()
internal_subset = internal_subset.rename(
    columns={internal_id_col: 'Internal_Student_ID', internal_mark_col: 'Internal_Mark'})

# Merge both sheets on Clean_ID
merged_df = pd.merge(bb_subset, internal_subset, on='Clean_ID', how='inner')

print(f"\n-> Total students successfully matched between both files: {len(merged_df)}")

if len(merged_df) > 0:
    print("\n--- DEBUG PREVIEW (First 5 Matched Students) ---")
    print(merged_df[['Clean_ID', 'BB_Score', 'Internal_Mark']].head())
    print("-" * 50)

# Convert scores to numeric for precise comparison
merged_df['BB_Score_Num'] = pd.to_numeric(merged_df['BB_Score'], errors='coerce')
merged_df['Internal_Mark_Num'] = pd.to_numeric(merged_df['Internal_Mark'], errors='coerce')

# Identify grade mismatches
mismatches = merged_df[
    (merged_df['BB_Score_Num'] != merged_df['Internal_Mark_Num']) &
    ~(merged_df['BB_Score_Num'].isna() & merged_df['Internal_Mark_Num'].isna())
    ].copy()

# ---------------------------------------------------------------------------
# 7. Display & Save Results
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print(f"Total grade mismatches found: {len(mismatches)}")
print("=" * 60)

if not mismatches.empty:
    display_cols = [c for c in ['BB_Student_ID', 'Last Name', 'First Name', 'BB_Score', 'Internal_Mark'] if
                    c in mismatches.columns]
    print("\nGrade Mismatches Found:")
    print(mismatches[display_cols])
else:
    print(
        "\nNo grade mismatches found! (Check the debug preview above to see if your scores are showing up as expected).")

# Save Excel report
output_filename = "grade_mismatches_report.xlsx"
export_df = mismatches.drop(columns=['Clean_ID', 'BB_Score_Num', 'Internal_Mark_Num'], errors='ignore')
export_df.to_excel(output_filename, index=False)

print(f"\nReport saved as '{output_filename}'.")