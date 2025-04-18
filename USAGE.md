# FinMod V2: Spreadsheet Processing Framework - Usage Guide

This guide provides instructions on how to use the Spreadsheet Processing Framework for compressing and encoding Excel spreadsheets.

## Installation

1. Clone the repository or download the source code
2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Basic Usage

Process a spreadsheet with default settings:
```bash
python main.py path/to/your/spreadsheet.xlsx
```

This will:
1. Parse the spreadsheet
2. Apply structural anchor-based compression
3. Encode the original and compressed versions
4. Save the results to the `output` directory

## Command Line Options

The script provides several options to customize processing:

```
python main.py path/to/your/spreadsheet.xlsx [OPTIONS]
```

Available options:
- `--output-dir DIR`: Specify output directory (default: "output")
- `--max-tokens N`: Maximum token limit for encoding (default: 4000)
- `--no-anchors`: Disable structural anchor-based extraction
- `--use-inverted-index`: Enable inverted-index translation (experimental)
- `--use-aggregation`: Enable data-format-aware aggregation (experimental)
- `--anchor-proximity N`: Set proximity range for structural anchors (default: 4)

## Output Files

For an input file `spreadsheet.xlsx`, the following output files are generated:

- `spreadsheet_original.md`: Markdown-like encoding of the original spreadsheet
- `spreadsheet_compressed.md`: Encoding of the compressed spreadsheet
- `spreadsheet_stats.json`: Compression statistics including token counts and reduction ratios

## Advanced Usage

### Disable All Compression

To encode a spreadsheet without any compression:
```bash
python main.py path/to/your/spreadsheet.xlsx --no-anchors
```

### Experimental Features

To enable experimental compression features:
```bash
python main.py path/to/your/spreadsheet.xlsx --use-inverted-index --use-aggregation
```

## Compression Strategies

The framework implements the following compression strategies:

1. **Structural Anchor-Based Extraction** (enabled by default)
   - Identifies "anchor" rows and columns (headers, data type transitions, etc.)
   - Retains cells within a proximity range of these anchors
   - Proximity range can be adjusted with `--anchor-proximity`

2. **Inverted-Index Translation** (experimental)
   - Groups repeated values and formats
   - Creates a more compact representation

3. **Data-Format-Aware Aggregation** (experimental)
   - Merges adjacent cells with similar data types
   - Creates a more compact representation

## Processing Large Spreadsheets

For very large spreadsheets, consider:
- Increasing the token limit: `--max-tokens 8000`
- Using a higher anchor proximity: `--anchor-proximity 6`
- Adding memory profiling: `python -m memory_profiler main.py path/to/spreadsheet.xlsx` 