#!/usr/bin/env python3
"""
Test script to verify the implementation of the Spreadsheet Processing Framework.
"""

import os
import sys
from pathlib import Path

from Core.parser import SpreadsheetParser
from Core.encoder import SpreadsheetEncoder
from Core.compressor import SheetCompressor
from Core.utils.helpers import estimate_tokens

def main():
    """Run basic tests for the framework."""
    # Check if sample data exists, generate if not
    sample_path = "Data/samples/sample_data.xlsx"
    if not os.path.exists(sample_path):
        print("Sample data not found, generating...")
        from generate_sample import generate_sample_data
        generate_sample_data(rows=50, cols=15, output_file=sample_path)
    
    print(f"Using sample data: {sample_path}")
    
    # Test parser
    print("\nTesting parser...")
    parser = SpreadsheetParser(sample_path)
    sheet_matrix = parser.parse_sheet()
    
    print(f"Sheet Name: {sheet_matrix.sheet_name}")
    print(f"Dimensions: {sheet_matrix.max_row}x{sheet_matrix.max_col}")
    
    # Count non-empty cells
    non_empty_count = sum(1 for row in range(1, sheet_matrix.max_row + 1) 
                         for col in range(1, sheet_matrix.max_col + 1) 
                         if sheet_matrix.get_cell(row, col) is not None and 
                            sheet_matrix.get_cell(row, col).value is not None)
    
    print(f"Non-empty cells: {non_empty_count}")
    
    # Test encoder
    print("\nTesting encoder...")
    encoder = SpreadsheetEncoder(max_tokens=8000)
    
    try:
        encoded = encoder.encode_sheet(sheet_matrix)
        token_count = estimate_tokens(encoded)
        print(f"Encoded with 8000 token limit: {token_count} tokens used")
        
        # Save encoded output for inspection
        output_dir = "output"
        os.makedirs(output_dir, exist_ok=True)
        with open(os.path.join(output_dir, "encoded_original.md"), "w") as f:
            f.write(encoded)
            
        print(f"Encoded output saved to output/encoded_original.md")
        
    except ValueError as e:
        print(f"Failed to encode: {str(e)}")
    
    # Test compressor
    print("\nTesting compressor...")
    compressor = SheetCompressor(
        use_structural_anchors=True,
        use_inverted_index=False,
        use_format_aggregation=False,
        anchor_proximity=4
    )
    
    compressed_matrix = compressor.compress(sheet_matrix)
    
    # Count cells after compression
    compressed_cell_count = sum(1 for row in range(1, compressed_matrix.max_row + 1) 
                               for col in range(1, compressed_matrix.max_col + 1) 
                               if compressed_matrix.get_cell(row, col) is not None and 
                                  compressed_matrix.get_cell(row, col).value is not None)
    
    print(f"Original cell count: {non_empty_count}")
    print(f"Compressed cell count: {compressed_cell_count}")
    print(f"Compression ratio: {compressed_cell_count / non_empty_count:.2f}")
    
    # Test encoding the compressed matrix
    try:
        encoded_compressed = encoder.encode_sheet(compressed_matrix)
        compressed_tokens = estimate_tokens(encoded_compressed)
        print(f"Compressed token count: {compressed_tokens}")
        
        # Save compressed output
        with open(os.path.join(output_dir, "encoded_compressed.md"), "w") as f:
            f.write(encoded_compressed)
            
        print(f"Compressed output saved to output/encoded_compressed.md")
        
    except ValueError as e:
        print(f"Failed to encode compressed sheet: {str(e)}")
    
    print("\nAll tests completed.")

if __name__ == "__main__":
    main() 