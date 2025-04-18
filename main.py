#!/usr/bin/env python3
"""
Main script to demonstrate spreadsheet processing framework.
"""

import os
import argparse
import tiktoken
import json
from pathlib import Path

from Core.parser import SpreadsheetParser
from Core.encoder import SpreadsheetEncoder
from Core.compressor import SheetCompressor
from Core.utils.helpers import estimate_tokens

def process_spreadsheet(filepath: str, output_dir: str, max_tokens: int = 4000, 
                        use_structural_anchors: bool = True, 
                        use_inverted_index: bool = False,
                        use_format_aggregation: bool = False,
                        anchor_proximity: int = 4):
    """Process a single spreadsheet file."""
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Parse the spreadsheet
    print(f"Parsing spreadsheet: {filepath}")
    parser = SpreadsheetParser(filepath)
    sheet_matrix = parser.parse_sheet()
    
    # Get statistics before compression
    original_cell_count = sum(1 for row in range(1, sheet_matrix.max_row + 1) 
                             for col in range(1, sheet_matrix.max_col + 1) 
                             if sheet_matrix.get_cell(row, col) is not None and 
                                sheet_matrix.get_cell(row, col).value is not None)
    
    print(f"Original dimensions: {sheet_matrix.max_row}x{sheet_matrix.max_col}")
    print(f"Original cell count: {original_cell_count}")
    
    # Encode without compression first to get baseline token count
    encoder = SpreadsheetEncoder(max_tokens=max_tokens)
    
    try:
        encoded_original = encoder.encode_sheet(sheet_matrix)
        original_tokens = estimate_tokens(encoded_original)
        print(f"Original token count: {original_tokens}")
        
        # Save original encoding
        original_filename = os.path.join(output_dir, f"{Path(filepath).stem}_original.md")
        with open(original_filename, 'w') as f:
            f.write(encoded_original)
            
        print(f"Original encoding saved to: {original_filename}")
    except ValueError as e:
        print(f"Warning: {str(e)}")
        print("Proceeding with compression...")
        original_tokens = max_tokens + 1  # Assume it's over the limit
    
    # Apply compression
    print("\nApplying compression strategies:")
    
    if use_structural_anchors:
        print(f"- Structural anchor-based extraction (proximity: {anchor_proximity})")
    if use_inverted_index:
        print("- Inverted-index translation")
    if use_format_aggregation:
        print("- Data-format-aware aggregation")
    
    compressor = SheetCompressor(
        use_structural_anchors=use_structural_anchors,
        use_inverted_index=use_inverted_index,
        use_format_aggregation=use_format_aggregation,
        anchor_proximity=anchor_proximity
    )
    
    compressed_matrix = compressor.compress(sheet_matrix)
    
    # Get statistics after compression
    compressed_cell_count = sum(1 for row in range(1, compressed_matrix.max_row + 1) 
                               for col in range(1, compressed_matrix.max_col + 1) 
                               if compressed_matrix.get_cell(row, col) is not None and 
                                  compressed_matrix.get_cell(row, col).value is not None)
    
    print(f"\nCompressed cell count: {compressed_cell_count}")
    print(f"Compression ratio: {compressed_cell_count / original_cell_count:.2f}")
    
    # Encode compressed sheet
    try:
        encoded_compressed = encoder.encode_sheet(compressed_matrix)
        compressed_tokens = estimate_tokens(encoded_compressed)
        print(f"Compressed token count: {compressed_tokens}")
        print(f"Token reduction: {original_tokens - compressed_tokens} ({(original_tokens - compressed_tokens) / original_tokens * 100:.2f}%)")
        
        # Save compressed encoding
        compressed_filename = os.path.join(output_dir, f"{Path(filepath).stem}_compressed.md")
        with open(compressed_filename, 'w') as f:
            f.write(encoded_compressed)
            
        print(f"Compressed encoding saved to: {compressed_filename}")
    except ValueError as e:
        print(f"Error encoding compressed sheet: {str(e)}")
        return
    
    # Save compression details
    stats = {
        "filename": filepath,
        "original_dimensions": f"{sheet_matrix.max_row}x{sheet_matrix.max_col}",
        "original_cell_count": original_cell_count,
        "original_tokens": original_tokens,
        "compressed_cell_count": compressed_cell_count,
        "compressed_tokens": compressed_tokens,
        "cell_compression_ratio": compressed_cell_count / original_cell_count,
        "token_reduction_ratio": compressed_tokens / original_tokens,
        "token_savings": original_tokens - compressed_tokens,
        "compression_strategies": {
            "structural_anchors": use_structural_anchors,
            "anchor_proximity": anchor_proximity,
            "inverted_index": use_inverted_index,
            "format_aggregation": use_format_aggregation
        }
    }
    
    stats_filename = os.path.join(output_dir, f"{Path(filepath).stem}_stats.json")
    with open(stats_filename, 'w') as f:
        json.dump(stats, f, indent=2)
        
    print(f"Compression statistics saved to: {stats_filename}")

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Spreadsheet Processing Framework")
    
    parser.add_argument("filepath", help="Path to the Excel spreadsheet file")
    parser.add_argument("--output-dir", default="output", help="Directory to save outputs")
    parser.add_argument("--max-tokens", type=int, default=4000, help="Maximum token limit")
    
    # Compression strategy toggles
    parser.add_argument("--no-anchors", action="store_true", help="Disable structural-anchor-based extraction")
    parser.add_argument("--use-inverted-index", action="store_true", help="Enable inverted-index translation")
    parser.add_argument("--use-aggregation", action="store_true", help="Enable data-format-aware aggregation")
    parser.add_argument("--anchor-proximity", type=int, default=4, help="Proximity range for structural anchors")
    
    args = parser.parse_args()
    
    process_spreadsheet(
        filepath=args.filepath,
        output_dir=args.output_dir,
        max_tokens=args.max_tokens,
        use_structural_anchors=not args.no_anchors,
        use_inverted_index=args.use_inverted_index,
        use_format_aggregation=args.use_aggregation,
        anchor_proximity=args.anchor_proximity
    )

if __name__ == "__main__":
    main() 