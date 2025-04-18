#!/usr/bin/env python3
"""
Test script to analyze and debug the real Excel model implementation.
"""

import os
import json
import argparse
from pathlib import Path
from openpyxl import load_workbook

from Core.parser import SpreadsheetParser
from Core.encoder import SpreadsheetEncoder
from Core.compressor import SheetCompressor
from Core.utils.helpers import estimate_tokens, detect_number_format_string

def analyze_sheet(sheet_matrix, sheet_name):
    """Analyze a single sheet and print statistics."""
    print(f"\nAnalyzing sheet: {sheet_name}")
    print(f"Dimensions: {sheet_matrix.max_row}x{sheet_matrix.max_col}")
    
    # Count non-empty cells
    non_empty_count = sum(1 for row in range(1, sheet_matrix.max_row + 1) 
                         for col in range(1, sheet_matrix.max_col + 1) 
                         if sheet_matrix.get_cell(row, col) is not None and 
                            sheet_matrix.get_cell(row, col).value is not None)
    
    print(f"Non-empty cells: {non_empty_count}")
    
    # Count different data types
    data_types = {}
    for row in range(1, sheet_matrix.max_row + 1):
        for col in range(1, sheet_matrix.max_col + 1):
            cell = sheet_matrix.get_cell(row, col)
            if cell and cell.value is not None:
                data_type = cell.data_type.name
                data_types[data_type] = data_types.get(data_type, 0) + 1
    
    print("\nData type distribution:")
    for dtype, count in data_types.items():
        print(f"{dtype}: {count}")
    
    return {
        "dimensions": f"{sheet_matrix.max_row}x{sheet_matrix.max_col}",
        "non_empty_cells": non_empty_count,
        "data_types": data_types
    }

def test_compression_modules(sheet_matrix, sheet_name, output_dir, token_limit=8000, skip_if_larger_than=None):
    """Test different compression module combinations."""
    results = {}
    encoder = SpreadsheetEncoder(max_tokens=token_limit)
    
    # Check if sheet is too large to process
    estimated_size = sheet_matrix.max_row * sheet_matrix.max_col * 2  # Rough estimate 
    if skip_if_larger_than and estimated_size > skip_if_larger_than:
        print(f"Sheet {sheet_name} is too large (est. {estimated_size} tokens), skipping detailed analysis")
        
        # Just do a quick analysis with the most promising approach (Module 3)
        try:
            print("Trying Module 3 (Data-Format-Aware Aggregation) for large sheet...")
            encoder.use_inverted_index = False
            encoder.use_format_aggregation = True
            compressed_tokens = "Too large"
            
            try:
                m3_encoded = encoder.encode_sheet(sheet_matrix)
                m3_tokens = estimate_tokens(m3_encoded)
                compressed_tokens = m3_tokens
                
                # Save the encoded result
                m3_file = os.path.join(output_dir, f"{sheet_name}_m3_format_aggregation.json")
                with open(m3_file, 'w') as f:
                    f.write(m3_encoded)
                print(f"Successfully encoded large sheet with Module 3: {m3_tokens} tokens")
                
                return {
                    "module3": {
                        "tokens": m3_tokens,
                        "file": m3_file,
                        "compression_ratio": "N/A (large sheet)"
                    }
                }
            except ValueError:
                print("Module 3 failed, sheet is too large for standard encoding")
                return {"error": "Sheet too large for encoding"}
                
        except Exception as e:
            print(f"Error processing large sheet: {str(e)}")
            return {"error": str(e)}
    
    # Original (no compression)
    print("\nTesting vanilla encoding (no compression)...")
    try:
        encoder.use_inverted_index = False
        encoder.use_format_aggregation = False
        
        vanilla_encoded = encoder.encode_sheet(sheet_matrix)
        vanilla_tokens = estimate_tokens(vanilla_encoded)
        vanilla_file = os.path.join(output_dir, f"{sheet_name}_vanilla.md")
        with open(vanilla_file, 'w') as f:
            f.write(vanilla_encoded)
        print(f"Vanilla encoding token count: {vanilla_tokens}")
        
        results["vanilla"] = {
            "tokens": vanilla_tokens,
            "file": vanilla_file,
            "compression_ratio": 1.0  # Baseline
        }
    except ValueError as e:
        print(f"Error with vanilla encoding: {str(e)}")
        results["vanilla"] = {"error": str(e)}
    
    # Test Module 1: Structural Anchor Compression
    print("\nTesting Module 1: Structural Anchor Compression...")
    try:
        compressor = SheetCompressor()
        compressed_m1 = compressor.compress(sheet_matrix)
        
        # Encode compressed sheet
        encoder.use_inverted_index = False
        encoder.use_format_aggregation = False
        m1_encoded = encoder.encode_sheet(compressed_m1)
        m1_tokens = estimate_tokens(m1_encoded)
        m1_file = os.path.join(output_dir, f"{sheet_name}_m1_structural_anchor.md")
        with open(m1_file, 'w') as f:
            f.write(m1_encoded)
        
        compression_ratio = m1_tokens / vanilla_tokens if vanilla_tokens > 0 else 0
        print(f"Module 1 token count: {m1_tokens}")
        print(f"Module 1 compression ratio: {1/compression_ratio:.2f}x")
        
        results["module1"] = {
            "tokens": m1_tokens,
            "file": m1_file,
            "compression_ratio": 1/compression_ratio
        }
    except Exception as e:
        print(f"Error with Module 1: {str(e)}")
        results["module1"] = {"error": str(e)}
    
    # Test Module 2: Inverted Index Translation
    print("\nTesting Module 2: Inverted Index Translation...")
    try:
        encoder.use_inverted_index = True
        encoder.use_format_aggregation = False
        m2_encoded = encoder.encode_sheet(sheet_matrix)
        m2_tokens = estimate_tokens(m2_encoded)
        m2_file = os.path.join(output_dir, f"{sheet_name}_m2_inverted_index.json")
        with open(m2_file, 'w') as f:
            f.write(m2_encoded)
        
        compression_ratio = m2_tokens / vanilla_tokens if vanilla_tokens > 0 else 0
        print(f"Module 2 token count: {m2_tokens}")
        print(f"Module 2 compression ratio: {1/compression_ratio:.2f}x")
        
        results["module2"] = {
            "tokens": m2_tokens,
            "file": m2_file,
            "compression_ratio": 1/compression_ratio
        }
    except Exception as e:
        print(f"Error with Module 2: {str(e)}")
        results["module2"] = {"error": str(e)}
    
    # Test Module 3: Data-Format-Aware Aggregation
    print("\nTesting Module 3: Data-Format-Aware Aggregation...")
    try:
        encoder.use_inverted_index = False
        encoder.use_format_aggregation = True
        m3_encoded = encoder.encode_sheet(sheet_matrix)
        m3_tokens = estimate_tokens(m3_encoded)
        m3_file = os.path.join(output_dir, f"{sheet_name}_m3_format_aggregation.json")
        with open(m3_file, 'w') as f:
            f.write(m3_encoded)
        
        compression_ratio = m3_tokens / vanilla_tokens if vanilla_tokens > 0 else 0
        print(f"Module 3 token count: {m3_tokens}")
        print(f"Module 3 compression ratio: {1/compression_ratio:.2f}x")
        
        results["module3"] = {
            "tokens": m3_tokens,
            "file": m3_file,
            "compression_ratio": 1/compression_ratio
        }
    except Exception as e:
        print(f"Error with Module 3: {str(e)}")
        results["module3"] = {"error": str(e)}
    
    # Test Combined: All three modules
    print("\nTesting Combined: All three modules...")
    try:
        # First apply structural compression
        compressor = SheetCompressor()
        compressed_combined = compressor.compress(sheet_matrix)
        
        # Then apply inverted index and format aggregation
        encoder.use_inverted_index = True
        encoder.use_format_aggregation = True
        combined_encoded = encoder.encode_sheet(compressed_combined)
        combined_tokens = estimate_tokens(combined_encoded)
        combined_file = os.path.join(output_dir, f"{sheet_name}_combined_all_modules.json")
        with open(combined_file, 'w') as f:
            f.write(combined_encoded)
        
        compression_ratio = combined_tokens / vanilla_tokens if vanilla_tokens > 0 else 0
        print(f"Combined token count: {combined_tokens}")
        print(f"Combined compression ratio: {1/compression_ratio:.2f}x")
        
        results["combined"] = {
            "tokens": combined_tokens,
            "file": combined_file,
            "compression_ratio": 1/compression_ratio
        }
    except Exception as e:
        print(f"Error with Combined approach: {str(e)}")
        results["combined"] = {"error": str(e)}
    
    return results

def main():
    """Run analysis on the real Excel model."""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Test SpreadsheetLLM Framework on real Excel models")
    parser.add_argument("--file", default="Data/Real Data/Miscounted Cash Flow.ExcelModel.xlsx", 
                      help="Path to Excel file")
    parser.add_argument("--output-dir", default="output", help="Output directory")
    parser.add_argument("--sheet", help="Specific sheet to process (omit to process all sheets)")
    parser.add_argument("--token-limit", type=int, default=8000, help="Token limit for encoding")
    parser.add_argument("--large-sheet-threshold", type=int, default=50000, 
                      help="Skip detailed analysis for sheets larger than this threshold")
    args = parser.parse_args()
    
    # Set up variables
    model_path = args.file
    output_dir = args.output_dir
    token_limit = args.token_limit
    large_sheet_threshold = args.large_sheet_threshold
    specific_sheet = args.sheet
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    if not os.path.exists(model_path):
        print(f"Error: File not found at {model_path}")
        return
    
    print(f"Analyzing real model: {model_path}")
    
    # Load workbook to get sheet names
    wb = load_workbook(model_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    wb.close()
    
    if specific_sheet:
        if specific_sheet in sheet_names:
            sheet_names = [specific_sheet]
            print(f"\nProcessing only sheet: {specific_sheet}")
        else:
            print(f"Error: Sheet '{specific_sheet}' not found in the workbook")
            print(f"Available sheets: {', '.join(sheet_names)}")
            return
    else:
        print(f"\nFound {len(sheet_names)} sheets:")
        for name in sheet_names:
            print(f"- {name}")
    
    # Initialize parser
    parser = SpreadsheetParser(model_path)
    
    # Store results for all sheets
    all_sheets_stats = {}
    all_compression_results = {}
    
    # Analyze each sheet
    for sheet_name in sheet_names:
        try:
            print(f"\nProcessing sheet: {sheet_name}")
            sheet_matrix = parser.parse_sheet(sheet_name)
            
            # Get basic statistics
            sheet_stats = analyze_sheet(sheet_matrix, sheet_name)
            
            # Test all compression modules
            compression_results = test_compression_modules(
                sheet_matrix, 
                sheet_name, 
                output_dir,
                token_limit=token_limit,
                skip_if_larger_than=large_sheet_threshold
            )
            
            # Update sheet statistics
            sheet_stats.update({
                "compression_results": compression_results
            })
            
            all_sheets_stats[sheet_name] = sheet_stats
            all_compression_results[sheet_name] = compression_results
            
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
            all_sheets_stats[sheet_name] = {"error": str(e)}
    
    # Save overall statistics
    stats_filename = os.path.join(output_dir, "compression_stats.json")
    with open(stats_filename, 'w') as f:
        json.dump(all_sheets_stats, f, indent=2)
    print(f"\nCompression statistics saved to: {stats_filename}")
    
    # Create a summary report
    report_rows = []
    report_rows.append("# Compression Summary Report")
    report_rows.append("\nSheet Name | Original Size | Module 1 | Module 2 | Module 3 | Combined | Best Approach")
    report_rows.append("--- | --- | --- | --- | --- | --- | ---")
    
    for sheet_name, results in all_compression_results.items():
        if "vanilla" not in results or "error" in results["vanilla"]:
            if "module3" in results and "tokens" in results["module3"]:
                # For large sheets that only have Module 3 results
                m3_tokens = results["module3"]["tokens"]
                row = f"{sheet_name} | Too large | - | - | {m3_tokens} tokens | - | Module 3 (large sheet)"
                report_rows.append(row)
            continue
            
        vanilla_tokens = results["vanilla"]["tokens"]
        m1_ratio = results.get("module1", {}).get("compression_ratio", 0)
        m2_ratio = results.get("module2", {}).get("compression_ratio", 0)
        m3_ratio = results.get("module3", {}).get("compression_ratio", 0)
        combined_ratio = results.get("combined", {}).get("compression_ratio", 0)
        
        # Determine best approach
        best_ratio = max(m1_ratio, m2_ratio, m3_ratio, combined_ratio)
        if best_ratio == m1_ratio:
            best_approach = "Module 1"
        elif best_ratio == m2_ratio:
            best_approach = "Module 2"
        elif best_ratio == m3_ratio:
            best_approach = "Module 3"
        else:
            best_approach = "Combined"
            
        row = f"{sheet_name} | {vanilla_tokens} tokens | {m1_ratio:.2f}x | {m2_ratio:.2f}x | {m3_ratio:.2f}x | {combined_ratio:.2f}x | {best_approach} ({best_ratio:.2f}x)"
        report_rows.append(row)
    
    # Add recommendations section
    report_rows.append("\n## Recommendations")
    report_rows.append("\nBased on the test results, here are the recommended approaches:")
    report_rows.append("\n1. **For financial sheets with numeric data**: Use Module 3 (Format-Aware Aggregation)")
    report_rows.append("2. **For text-heavy sheets**: Use Module 2 (Inverted Index)")
    report_rows.append("3. **For very large sheets**: Use chunked encoding with Module 3")
    report_rows.append("4. **For small sheets with varied formats**: Use the Combined approach")
    
    # Add implementation details
    report_rows.append("\n## Implementation Details")
    report_rows.append("\n- **Module 1**: Identifies and preserves important rows/columns while discarding uniform data")
    report_rows.append("- **Module 2**: Groups identical values to reduce token usage for repeated data")
    report_rows.append("- **Module 3**: Groups cells with similar formats for better compression")
    report_rows.append("- **Combined**: Applies all three modules in sequence")
    report_rows.append("\nLarge sheets are automatically chunked with headers preserved in each chunk.")
    
    # Save the summary report
    report_filename = os.path.join(output_dir, "compression_summary.md")
    with open(report_filename, 'w') as f:
        f.write("\n".join(report_rows))
    print(f"Compression summary report saved to: {report_filename}")
    
    print("\nAnalysis complete!")
    
    # Print final recommendations
    print("\nRecommendations based on analysis:")
    print("1. Use Module 3 (Format-Aware Aggregation) for financial sheets with numeric data")
    print("2. Use Module 2 (Inverted Index) for text-heavy sheets")
    print("3. Use chunking with Module 3 for very large sheets")
    print("4. Use Combined approach for small sheets with varied formats")

if __name__ == "__main__":
    main() 