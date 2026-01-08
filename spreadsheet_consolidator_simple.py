#!/usr/bin/env python3
"""
Spreadsheet Consolidator (Simple Version)
Merges data from two spreadsheets and applies merge & center formatting.

HOW TO USE:
1. Edit the settings below (COLUMNS_TO_EXTRACT, FILE_1, FILE_2, OUTPUT_FILE)
2. Run: python spreadsheet_consolidator_simple.py
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from pathlib import Path

# =============================================================================
# ⬇️ EDIT THESE SETTINGS ⬇️
# =============================================================================

# Your input files (put them in the same folder as this script, or use full paths)
FILE_1 = "corporate_actions.xlsx"      # ← Change to your first file name
FILE_2 = "accounts_list.xlsx"          # ← Change to your second file name

# Output file name
OUTPUT_FILE = "consolidated_output.xlsx"  # ← Change if you want a different name

# Columns to extract (in order from left to right)
# Column matching is case-insensitive: "SEDOL" matches "Sedol", "sedol", etc.
COLUMNS_TO_EXTRACT = [
    "ISSUE NAME",
    "SEDOL",
    "ISIN",
    "ISSUE COUNTRY NAME",
    "STRATEGY",
    "SUB STRATEGY",
    "FUND TYPE",
    "FUND NAME",
]

# Set to False if you don't want cells merged
APPLY_MERGE_FORMATTING = True

# =============================================================================
# ⬆️ EDIT ABOVE THIS LINE ⬆️
# =============================================================================


def normalize_column_name(name):
    """Convert column name to uppercase for matching."""
    if name is None:
        return ''
    return str(name).strip().upper()


def find_matching_column(df, target_col):
    """Find column in dataframe (case-insensitive)."""
    target_normalized = normalize_column_name(target_col)
    for col in df.columns:
        if normalize_column_name(col) == target_normalized:
            return col
    return None


def extract_columns(df, columns_to_extract):
    """Extract specified columns from dataframe."""
    extracted_data = {}
    for target_col in columns_to_extract:
        actual_col = find_matching_column(df, target_col)
        if actual_col is not None:
            normalized_name = normalize_column_name(target_col)
            extracted_data[normalized_name] = df[actual_col].values
            print(f"  ✓ Found '{actual_col}' -> '{normalized_name}'")
        else:
            print(f"  ✗ Column '{target_col}' not found")
    return pd.DataFrame(extracted_data)


def apply_hierarchical_merge(ws, start_row, end_row, columns):
    """Apply hierarchical merge and center formatting."""
    group_boundaries = [(start_row, end_row)]
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for col_idx in columns:
        col_letter = get_column_letter(col_idx)
        merge_ranges = []
        new_boundaries = []
        
        for group_start, group_end in group_boundaries:
            if group_start > group_end:
                continue
            
            current_value = ws.cell(row=group_start, column=col_idx).value
            range_start = group_start
            
            for row in range(group_start, group_end + 1):
                cell_value = ws.cell(row=row, column=col_idx).value
                
                if cell_value != current_value:
                    if row - 1 >= range_start:
                        if row - 1 > range_start:
                            merge_ranges.append((range_start, row - 1))
                        new_boundaries.append((range_start, row - 1))
                    current_value = cell_value
                    range_start = row
            
            if range_start <= group_end:
                if group_end > range_start:
                    merge_ranges.append((range_start, group_end))
                new_boundaries.append((range_start, group_end))
        
        group_boundaries = new_boundaries
        
        for start, end in merge_ranges:
            try:
                ws.merge_cells(f"{col_letter}{start}:{col_letter}{end}")
                ws.cell(row=start, column=col_idx).alignment = center_alignment
            except:
                pass
        
        for row in range(start_row, end_row + 1):
            try:
                ws.cell(row=row, column=col_idx).alignment = center_alignment
            except:
                pass


def apply_formatting(ws):
    """Apply header and cell formatting."""
    # Header formatting
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    ws.row_dimensions[1].height = 30
    
    # Data cell borders
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).border = thin_border
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 50)


def main():
    print("=" * 60)
    print("SPREADSHEET CONSOLIDATOR")
    print("=" * 60)
    
    # Check files exist
    if not Path(FILE_1).exists():
        print(f"\n❌ ERROR: Cannot find file '{FILE_1}'")
        print("   Make sure the file is in the same folder as this script,")
        print("   or update FILE_1 with the full path.")
        return
    
    if not Path(FILE_2).exists():
        print(f"\n❌ ERROR: Cannot find file '{FILE_2}'")
        print("   Make sure the file is in the same folder as this script,")
        print("   or update FILE_2 with the full path.")
        return
    
    # Load spreadsheets
    print(f"\n📂 Loading '{FILE_1}'...")
    df1 = pd.read_excel(FILE_1)
    print(f"   Found {len(df1)} rows, {len(df1.columns)} columns")
    
    print(f"\n📂 Loading '{FILE_2}'...")
    df2 = pd.read_excel(FILE_2)
    print(f"   Found {len(df2)} rows, {len(df2.columns)} columns")
    
    # Extract columns
    print(f"\n🔍 Extracting columns from file 1...")
    extracted1 = extract_columns(df1, COLUMNS_TO_EXTRACT)
    
    print(f"\n🔍 Extracting columns from file 2...")
    extracted2 = extract_columns(df2, COLUMNS_TO_EXTRACT)
    
    # Ensure same columns in both
    all_columns = [normalize_column_name(c) for c in COLUMNS_TO_EXTRACT]
    for col in all_columns:
        if col not in extracted1.columns:
            extracted1[col] = pd.NA
        if col not in extracted2.columns:
            extracted2[col] = pd.NA
    
    extracted1 = extracted1[all_columns]
    extracted2 = extracted2[all_columns]
    
    # Merge data
    print(f"\n🔗 Merging spreadsheets...")
    merged_df = pd.concat([extracted1, extracted2], ignore_index=True)
    merged_df = merged_df.dropna(how='all')
    print(f"   Combined: {len(merged_df)} rows")
    
    # Sort by all columns
    print(f"\n📊 Sorting data...")
    merged_df = merged_df.sort_values(by=all_columns, na_position='last')
    merged_df = merged_df.reset_index(drop=True)
    
    # Save to Excel
    print(f"\n💾 Saving to '{OUTPUT_FILE}'...")
    merged_df.to_excel(OUTPUT_FILE, index=False, sheet_name='Consolidated Data')
    
    # Apply formatting
    print(f"\n🎨 Applying formatting...")
    wb = openpyxl.load_workbook(OUTPUT_FILE)
    ws = wb.active
    
    apply_formatting(ws)
    
    if APPLY_MERGE_FORMATTING and ws.max_row > 1:
        print(f"   Merging duplicate cells...")
        apply_hierarchical_merge(ws, 2, ws.max_row, list(range(1, ws.max_column + 1)))
    
    wb.save(OUTPUT_FILE)
    
    print("\n" + "=" * 60)
    print(f"✅ DONE! Output saved to: {OUTPUT_FILE}")
    print("=" * 60)


if __name__ == "__main__":
    main()
