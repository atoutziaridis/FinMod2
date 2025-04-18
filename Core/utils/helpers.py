"""Helper functions for the Spreadsheet Processing Framework."""

import re
import datetime
import tiktoken
from typing import Any, Dict, List, Tuple, Union, Optional

from Core.utils.constants import DataType

def get_column_letter(col_idx: int) -> str:
    """Convert column index to letter (1 -> A, 27 -> AA)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def get_column_index(col_letter: str) -> int:
    """Convert column letter to index (A -> 1, AA -> 27)."""
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char.upper()) - 64)
    return result

def get_cell_address(row: int, col: int) -> str:
    """Convert row and column indices to cell address (1,1 -> A1)."""
    return f"{get_column_letter(col)}{row}"

def parse_cell_address(address: str) -> Tuple[int, int]:
    """Parse cell address to row and column indices (A1 -> 1,1)."""
    match = re.match(r'([A-Z]+)(\d+)', address)
    if not match:
        raise ValueError(f"Invalid cell address: {address}")
    col_letter, row = match.groups()
    return int(row), get_column_index(col_letter)

def infer_data_type(value: Any) -> DataType:
    """Infer data type from a cell value."""
    if value is None:
        return DataType.EMPTY
    
    if isinstance(value, bool):
        return DataType.BOOLEAN
        
    if isinstance(value, int):
        return DataType.INT_NUM
        
    if isinstance(value, float):
        return DataType.FLOAT_NUM
    
    # Handle string values
    if isinstance(value, str):
        # Check if the string is a number
        value_str = value.strip()
        
        # Empty or whitespace-only string
        if not value_str:
            return DataType.EMPTY
            
        # Check for percentage format
        if value_str.endswith('%') and _is_numeric(value_str[:-1]):
            return DataType.PERCENTAGE
            
        # Check for currency format
        if (value_str.startswith('$') or value_str.startswith('€') or 
            value_str.startswith('£')) and _is_numeric(value_str[1:]):
            return DataType.CURRENCY
        
        # Check for date format - simple patterns
        date_patterns = [
            r'^\d{4}-\d{2}-\d{2}$',  # YYYY-MM-DD
            r'^\d{1,2}/\d{1,2}/\d{2,4}$',  # M/D/YY or M/D/YYYY
            r'^\d{1,2}-\d{1,2}-\d{2,4}$',  # M-D-YY or M-D-YYYY
        ]
        for pattern in date_patterns:
            if re.match(pattern, value_str):
                return DataType.DATE
        
        # Check if it's a numeric string
        if _is_numeric(value_str):
            if '.' in value_str:
                return DataType.FLOAT_NUM
            else:
                return DataType.INT_NUM
    
    # Default to text for anything else
    return DataType.TEXT

def _is_numeric(value_str: str) -> bool:
    """Check if a string represents a numeric value."""
    try:
        float(value_str.replace(',', ''))
        return True
    except ValueError:
        return False

def estimate_tokens(text: str) -> int:
    """Estimate number of tokens in text using tiktoken."""
    try:
        encoding = tiktoken.get_encoding("cl100k_base")
        return len(encoding.encode(text))
    except:
        # Fallback estimation if tiktoken is not available
        # This is a very rough estimate
        return len(text) // 4

def detect_number_format_string(value: Any) -> Optional[str]:
    """Detect number format string based on value."""
    if value is None:
        return None
        
    if isinstance(value, bool):
        return None
        
    if isinstance(value, int):
        return "0"
        
    if isinstance(value, float):
        # Check if it's a whole number
        if value.is_integer():
            return "0"
            
        # Check decimal places
        str_value = str(value)
        if '.' in str_value:
            decimal_places = len(str_value.split('.')[1])
            return "0." + "0" * decimal_places
    
    # String checks
    if isinstance(value, str):
        value_str = value.strip()
        
        # Percentage
        if value_str.endswith('%') and _is_numeric(value_str[:-1]):
            return "0%"
            
        # Currency with 2 decimal places
        if (value_str.startswith('$') or value_str.startswith('€') or 
            value_str.startswith('£')) and _is_numeric(value_str[1:]):
            
            # Check for thousands separator
            if ',' in value_str:
                if '.' in value_str:
                    # Currency with decimals and thousands separator
                    return "$#,##0.00"
                else:
                    # Currency without decimals but with thousands separator
                    return "$#,##0"
            else:
                if '.' in value_str:
                    # Currency with decimals, no thousands separator
                    return "$0.00"
                else:
                    # Currency without decimals, no thousands separator
                    return "$0"
                    
        # Numbers with thousands separator
        if ',' in value_str and _is_numeric(value_str.replace(',', '')):
            if '.' in value_str:
                # With decimals
                decimal_places = len(value_str.split('.')[1])
                return "#,##0." + "0" * decimal_places
            else:
                # Without decimals
                return "#,##0"
                
    return None 