#!/usr/bin/env python3
"""
Main script to demonstrate spreadsheet processing framework.
"""

import os
import argparse
import json
from pathlib import Path

from Core.parser import SpreadsheetParser
from Core.compressor import SheetCompressor, compress_spreadsheet_to_file

def process_spreadsheet(filepath: str, output_dir: str, max_tokens: int = 4000, 
                        use_structural_anchors: bool = True, 
                        use_inverted_index: bool = True,
                        use_format_aggregation: bool = True,
                        anchor_proximity: int = 4):
    """Process a single spreadsheet file."""
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Parse the spreadsheet
    print(f"Parsing spreadsheet: {filepath}")
    parser = SpreadsheetParser(filepath)
    
    # Get the list of sheet names in the workbook
    sheet_names = parser.workbook.sheetnames
    
    # Process each sheet
    all_stats = []
    for sheet_name in sheet_names:
        print(f"\nProcessing sheet: {sheet_name}")
        sheet_matrix = parser.parse_sheet(sheet_name)
        
        # Get base filename from the original file
        base_filename = Path(filepath).stem
        
        # Create a filename prefix that includes the sheet name
        sheet_prefix = f"{base_filename}_{sheet_name.replace(' ', '_')}"
        
        # Apply compression and save the sheet
        try:
            stats = compress_spreadsheet_to_file(
                sheet_matrix=sheet_matrix,
                output_dir=output_dir,
                filename_prefix=sheet_prefix,
                use_structural_anchors=use_structural_anchors,
                use_inverted_index=use_inverted_index,
                use_format_aggregation=use_format_aggregation,
                anchor_proximity=anchor_proximity,
                max_tokens=max_tokens
            )
            
            print(f"Original dimensions: {stats['original_dimensions']}")
            print(f"Original cell count: {stats['original_cell_count']}")
            print(f"Compressed cell count: {stats['compressed_cell_count']}")
            print(f"Compression ratio: {stats['compression_ratio']:.2f}")
            print(f"Token count: {stats['token_count']}")
            print(f"Output saved to: {stats['output_file']}")
            
            all_stats.append(stats)
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
    
    # Save overall statistics
    stats_filename = os.path.join(output_dir, f"{Path(filepath).stem}_all_sheets_stats.json")
    with open(stats_filename, 'w') as f:
        json.dump(all_stats, f, indent=2)
        
    print(f"\nAll sheets processed. Statistics saved to: {stats_filename}")

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Spreadsheet Processing Framework")
    
    parser.add_argument("filepath", help="Path to the Excel spreadsheet file")
    parser.add_argument("--output-dir", default="output", help="Directory to save outputs")
    parser.add_argument("--max-tokens", type=int, default=4000, help="Maximum token limit")
    
    # Compression strategy toggles
    parser.add_argument("--no-anchors", action="store_true", help="Disable structural-anchor-based extraction")
    parser.add_argument("--use-inverted-index", action="store_true", help="Enable inverted-index translation")
    parser.add_argument("--use-format-aggregation", action="store_true", help="Enable data-format-aware aggregation")
    parser.add_argument("--anchor-proximity", type=int, default=4, help="Proximity range for structural anchors")
    
    args = parser.parse_args()
    
    process_spreadsheet(
        filepath=args.filepath,
        output_dir=args.output_dir,
        max_tokens=args.max_tokens,
        use_structural_anchors=not args.no_anchors,
        use_inverted_index=args.use_inverted_index,
        use_format_aggregation=args.use_format_aggregation,
        anchor_proximity=args.anchor_proximity
    )

if __name__ == "__main__":
    main() 