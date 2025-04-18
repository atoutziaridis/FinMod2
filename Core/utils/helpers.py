"""Helper functions for the Spreadsheet Processing Framework."""

import re
import datetime
import tiktoken
from typing import Any, Dict, List, Tuple, Union

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
    """Infer data type from cell value."""
    if value is None:
        return DataType.EMPTY
    
    if isinstance(value, bool):
        return DataType.BOOLEAN
    
    if isinstance(value, int):
        return DataType.INT_NUM
    
    if isinstance(value, float):
        return DataType.FLOAT_NUM
    
    if isinstance(value, str):
        # Check if string is a date/time
        if re.match(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', value):
            return DataType.DATE
        
        # Check if string is an email
        if re.match(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', value):
            return DataType.EMAIL
        
        # Check if string is a URL
        if re.match(r'^(http|https)://', value) or re.match(r'^www\.', value):
            return DataType.URL
        
        # Check if string is a percentage
        if re.match(r'^-?\d+(\.\d+)?%$', value):
            return DataType.PERCENTAGE
        
        # Check if string is a formula
        if value.startswith('='):
            return DataType.FORMULA
        
        # Check if string is a phone number
        if re.match(r'^\+?[\d\s-]{10,15}$', value):
            return DataType.PHONE
        
        # Default to text
        return DataType.TEXT
    
    if isinstance(value, datetime.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return DataType.DATE
        elif value.year == 1900 and value.month == 1 and value.day == 1:
            return DataType.TIME
        else:
            return DataType.DATETIME
    
    return DataType.TEXT

def estimate_tokens(text: str) -> int:
    """Estimate token count for a given text using tiktoken."""
    encoding = tiktoken.get_encoding("cl100k_base")  # OpenAI's encoding
    return len(encoding.encode(text)) 