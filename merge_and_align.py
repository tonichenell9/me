#!/usr/bin/env python3
"""
Excel Merge and Align Utility
Merges data from two Excel spreadsheets and applies merge & centre formatting
to consecutive cells with identical values in specified columns.
"""

import logging
from pathlib import Path
from typing import Optional, List, Union
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def merge_excel_files(
    file1_path: Union[str, Path],
    file2_path: Union[str, Path],
    output_path: Union[str, Path],
    merge_type: str = 'vertical',
    sheet1_name: Optional[str] = None,
    sheet2_name: Optional[str] = None
) -> Workbook:
    """
    Merge data from two Excel files into a single workbook.
    
    Args:
        file1_path: Path to the first Excel file
        file2_path: Path to the second Excel file
        output_path: Path for the output merged file
        merge_type: 'vertical' (stack rows) or 'horizontal' (add columns side by side)
        sheet1_name: Optional name of sheet to read from file1 (defaults to active sheet)
        sheet2_name: Optional name of sheet to read from file2 (defaults to active sheet)
    
    Returns:
        The merged openpyxl Workbook object
    """
    file1_path = Path(file1_path)
    file2_path = Path(file2_path)
    output_path = Path(output_path)
    
    logger.info(f"Loading first file: {file1_path}")
    df1 = pd.read_excel(file1_path, sheet_name=sheet1_name or 0)
    
    logger.info(f"Loading second file: {file2_path}")
    df2 = pd.read_excel(file2_path, sheet_name=sheet2_name or 0)
    
    if merge_type == 'vertical':
        logger.info("Merging files vertically (stacking rows)")
        merged_df = pd.concat([df1, df2], ignore_index=True)
    elif merge_type == 'horizontal':
        logger.info("Merging files horizontally (side by side)")
        merged_df = pd.concat([df1, df2], axis=1)
    else:
        raise ValueError(f"Invalid merge_type: {merge_type}. Use 'vertical' or 'horizontal'")
    
    logger.info(f"Merged data shape: {merged_df.shape}")
    
    # Save to Excel using pandas first
    merged_df.to_excel(output_path, index=False)
    
    # Load with openpyxl for formatting
    workbook = load_workbook(output_path)
    logger.info(f"Merged file saved to: {output_path}")
    
    return workbook


def merge_and_centre_cells(
    workbook: Workbook,
    columns_to_merge: Optional[List[Union[int, str]]] = None,
    sheet_name: Optional[str] = None,
    start_row: int = 2,  # Default to 2 to skip header row
    end_row: Optional[int] = None
) -> Workbook:
    """
    Merge and centre consecutive cells with identical values in specified columns.
    
    Args:
        workbook: The openpyxl Workbook to process
        columns_to_merge: List of column indices (1-based) or letters to merge.
                         If None, all columns will be processed.
        sheet_name: Name of the sheet to process. Defaults to active sheet.
        start_row: Row to start merging from (1-based, default 2 to skip header)
        end_row: Row to end merging at. Defaults to last row with data.
    
    Returns:
        The modified Workbook object
    """
    if sheet_name:
        ws = workbook[sheet_name]
    else:
        ws = workbook.active
    
    if end_row is None:
        end_row = ws.max_row
    
    # Convert column references to indices if needed
    if columns_to_merge is None:
        columns_to_merge = list(range(1, ws.max_column + 1))
    else:
        processed_columns = []
        for col in columns_to_merge:
            if isinstance(col, str):
                # Convert column letter to index
                col_idx = openpyxl.utils.column_index_from_string(col)
                processed_columns.append(col_idx)
            else:
                processed_columns.append(col)
        columns_to_merge = processed_columns
    
    logger.info(f"Processing columns: {columns_to_merge}")
    logger.info(f"Row range: {start_row} to {end_row}")
    
    # Process each column
    for col_idx in columns_to_merge:
        col_letter = get_column_letter(col_idx)
        logger.info(f"Processing column {col_letter} (index {col_idx})")
        
        # Find consecutive cells with same values
        merge_ranges = []
        current_value = None
        range_start = start_row
        
        for row in range(start_row, end_row + 1):
            cell_value = ws.cell(row=row, column=col_idx).value
            
            if row == start_row:
                current_value = cell_value
                range_start = row
            elif cell_value != current_value:
                # End of current range
                if row - 1 > range_start:
                    merge_ranges.append((range_start, row - 1))
                current_value = cell_value
                range_start = row
        
        # Handle the last range
        if end_row > range_start and ws.cell(row=end_row, column=col_idx).value == current_value:
            merge_ranges.append((range_start, end_row))
        
        # Apply merge and centre to identified ranges
        for start, end in merge_ranges:
            merge_ref = f"{col_letter}{start}:{col_letter}{end}"
            logger.info(f"  Merging cells: {merge_ref}")
            
            try:
                ws.merge_cells(merge_ref)
                # Apply centre alignment to the merged cell
                merged_cell = ws.cell(row=start, column=col_idx)
                merged_cell.alignment = Alignment(
                    horizontal='center',
                    vertical='center',
                    wrap_text=True
                )
            except Exception as e:
                logger.warning(f"  Could not merge {merge_ref}: {e}")
    
    return workbook


def process_excel_merge_and_align(
    file1_path: Union[str, Path],
    file2_path: Union[str, Path],
    output_path: Union[str, Path],
    merge_type: str = 'vertical',
    columns_to_merge: Optional[List[Union[int, str]]] = None,
    skip_header: bool = True,
    sheet1_name: Optional[str] = None,
    sheet2_name: Optional[str] = None
) -> None:
    """
    Complete pipeline: merge two Excel files and apply merge & centre formatting.
    
    Args:
        file1_path: Path to the first Excel file
        file2_path: Path to the second Excel file  
        output_path: Path for the output file
        merge_type: 'vertical' (stack rows) or 'horizontal' (add columns)
        columns_to_merge: List of columns to apply merge & centre. 
                         Can be integers (1-based) or letters (e.g., 'A', 'B').
                         If None, all columns are processed.
        skip_header: If True, starts merging from row 2 (skips header row)
        sheet1_name: Name of sheet to read from file1
        sheet2_name: Name of sheet to read from file2
    """
    logger.info("=" * 60)
    logger.info("Starting Excel Merge and Align Process")
    logger.info("=" * 60)
    
    # Step 1: Merge the files
    workbook = merge_excel_files(
        file1_path=file1_path,
        file2_path=file2_path,
        output_path=output_path,
        merge_type=merge_type,
        sheet1_name=sheet1_name,
        sheet2_name=sheet2_name
    )
    
    # Step 2: Apply merge and centre formatting
    start_row = 2 if skip_header else 1
    workbook = merge_and_centre_cells(
        workbook=workbook,
        columns_to_merge=columns_to_merge,
        start_row=start_row
    )
    
    # Step 3: Save the final workbook
    output_path = Path(output_path)
    workbook.save(output_path)
    logger.info(f"Final output saved to: {output_path}")
    logger.info("Process completed successfully!")


def apply_merge_centre_to_existing(
    input_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    columns_to_merge: Optional[List[Union[int, str]]] = None,
    sheet_name: Optional[str] = None,
    skip_header: bool = True
) -> None:
    """
    Apply merge and centre formatting to an existing Excel file.
    
    Args:
        input_path: Path to the Excel file to process
        output_path: Path for the output file. If None, overwrites input file.
        columns_to_merge: List of columns to apply merge & centre.
                         Can be integers (1-based) or letters (e.g., 'A', 'B').
                         If None, all columns are processed.
        sheet_name: Name of sheet to process. Defaults to active sheet.
        skip_header: If True, starts merging from row 2 (skips header row)
    """
    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path
    else:
        output_path = Path(output_path)
    
    logger.info(f"Loading workbook: {input_path}")
    workbook = load_workbook(input_path)
    
    start_row = 2 if skip_header else 1
    workbook = merge_and_centre_cells(
        workbook=workbook,
        columns_to_merge=columns_to_merge,
        sheet_name=sheet_name,
        start_row=start_row
    )
    
    workbook.save(output_path)
    logger.info(f"Output saved to: {output_path}")


# Example usage and CLI
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Merge Excel files and apply merge & centre formatting to consecutive identical values"
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Subcommand: merge - Merge two files and apply formatting
    merge_parser = subparsers.add_parser('merge', help='Merge two Excel files')
    merge_parser.add_argument('file1', help='Path to first Excel file')
    merge_parser.add_argument('file2', help='Path to second Excel file')
    merge_parser.add_argument('output', help='Path for output file')
    merge_parser.add_argument(
        '--merge-type', '-t',
        choices=['vertical', 'horizontal'],
        default='vertical',
        help='How to merge: vertical (stack rows) or horizontal (side by side)'
    )
    merge_parser.add_argument(
        '--columns', '-c',
        nargs='*',
        help='Columns to merge & centre (e.g., A B C or 1 2 3). If not specified, all columns.'
    )
    merge_parser.add_argument(
        '--no-skip-header',
        action='store_true',
        help='Include header row in merge operations'
    )
    
    # Subcommand: format - Apply merge & centre to existing file
    format_parser = subparsers.add_parser('format', help='Apply merge & centre to existing file')
    format_parser.add_argument('input', help='Path to Excel file')
    format_parser.add_argument('--output', '-o', help='Output path (default: overwrite input)')
    format_parser.add_argument(
        '--columns', '-c',
        nargs='*',
        help='Columns to merge & centre (e.g., A B C or 1 2 3). If not specified, all columns.'
    )
    format_parser.add_argument(
        '--sheet', '-s',
        help='Sheet name to process (default: active sheet)'
    )
    format_parser.add_argument(
        '--no-skip-header',
        action='store_true',
        help='Include header row in merge operations'
    )
    
    args = parser.parse_args()
    
    # Parse columns argument
    def parse_columns(cols):
        if cols is None:
            return None
        parsed = []
        for c in cols:
            try:
                parsed.append(int(c))
            except ValueError:
                parsed.append(c.upper())
        return parsed
    
    if args.command == 'merge':
        columns = parse_columns(args.columns)
        process_excel_merge_and_align(
            file1_path=args.file1,
            file2_path=args.file2,
            output_path=args.output,
            merge_type=args.merge_type,
            columns_to_merge=columns,
            skip_header=not args.no_skip_header
        )
    
    elif args.command == 'format':
        columns = parse_columns(args.columns)
        apply_merge_centre_to_existing(
            input_path=args.input,
            output_path=args.output,
            columns_to_merge=columns,
            sheet_name=args.sheet,
            skip_header=not args.no_skip_header
        )
    
    else:
        parser.print_help()
        print("\n" + "=" * 60)
        print("EXAMPLES:")
        print("=" * 60)
        print("\n1. Merge two files vertically and apply merge & centre to all columns:")
        print("   python merge_and_align.py merge file1.xlsx file2.xlsx output.xlsx")
        print("\n2. Merge horizontally and only merge & centre columns A and B:")
        print("   python merge_and_align.py merge file1.xlsx file2.xlsx output.xlsx -t horizontal -c A B")
        print("\n3. Apply merge & centre to an existing file (columns 1, 2, 3):")
        print("   python merge_and_align.py format input.xlsx -c 1 2 3")
        print("\n4. Apply merge & centre to specific sheet:")
        print("   python merge_and_align.py format input.xlsx -o output.xlsx --sheet Sheet1 -c A B C")
