# FinMod V2: Spreadsheet Processing Framework

A modular preprocessing framework for transforming raw spreadsheets into compact, token-efficient encodings while preserving structural integrity.

## Features

- Spreadsheet parsing and structural preservation
- Data-type and format inference
- Efficient encoding for token reduction
- Compression modules with structural anchors, inverted indexing, and format-aware aggregation
- Table detection and QA capabilities

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