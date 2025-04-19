# FinMod V2 - Financial Modeling Framework

A framework for processing, compressing, and working with financial spreadsheets.

## Overview

This framework provides tools to:
1. Parse Excel spreadsheets
2. Compress spreadsheet data while preserving essential information
3. Encode spreadsheet data into a token-efficient format
4. Save only the final combined output for each sheet

## Features

- **Structural anchor-based extraction**: Identifies important rows and columns based on content patterns and formatting
- **Inverted index translation**: Efficiently represents cell values by mapping values to cell addresses
- **Format aggregation**: Groups cells with similar formatting to reduce redundancy
- **Sheet-by-sheet processing**: Process all sheets in a workbook individually
- **Final output only**: Only saves the final combined output for each sheet without intermediate files

## Usage

Basic usage:

```python
from Core import SpreadsheetParser, compress_spreadsheet_to_file

# Parse a spreadsheet
parser = SpreadsheetParser("path/to/spreadsheet.xlsx")
sheet_matrix = parser.parse_sheet("Sheet1")

# Compress and save
stats = compress_spreadsheet_to_file(
    sheet_matrix=sheet_matrix,
    output_dir="output",
    filename_prefix="my_sheet"
)

print(f"Compression ratio: {stats['compression_ratio']:.2f}")
print(f"Output saved to: {stats['output_file']}")
```

Or use the command-line interface:

```bash
python main.py path/to/spreadsheet.xlsx --output-dir output
```

See `USAGE.md` for more detailed instructions.

## Installation

```bash
pip install -r requirements.txt
```

## Project Structure

- `core/`: Core implementation modules
  - `parser.py`: Spreadsheet parsing logic
  - `encoder.py`: Data encoding utilities
  - `compressor.py`: Compression modules
  - `utils.py`: Utility functions
- `data/`: Sample data and test files 