"""
Command-line runner for Cut Planning Model v2, Page 1 (no web server needed).

Usage:
    python run_cli.py path/to/input.xlsx path/to/output.xlsx
    python run_cli.py                      # uses bundled sample data
"""

import os
import sys

from cutplan2 import run

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT = os.path.join(BASE_DIR, "sample_data", "BufferCuttingOrderForm_2026-08-21.xlsx")
DEFAULT_OUTPUT = os.path.join(BASE_DIR, "outputs", "Extracted_Data.xlsx")

if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    output_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df, missing_columns, filtered_out_count = run(input_path, output_path)

    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print(f"Extracted {len(df)} row(s), {len(df.columns)} field(s).")
    if filtered_out_count:
        print(f"Filtered out {filtered_out_count} row(s) whose Sewing Line did not start with 'VS'.")
    if missing_columns:
        print(f"Fields not found in input (left blank): {missing_columns}")
    else:
        print(f"All {len(df.columns)} fields were found in the input file.")
    print()
    print(df.head(20).to_string(index=False))
