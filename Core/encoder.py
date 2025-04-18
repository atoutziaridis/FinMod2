"""Encoder module for serializing spreadsheet data into token-efficient format."""

from typing import Dict, List, Tuple, Any, Optional, Set, Union
from collections import defaultdict
import json

from Core.parser import SheetMatrix, Cell
from Core.utils.constants import DataType, FormatType
from Core.utils.helpers import estimate_tokens, get_cell_address

class SpreadsheetEncoder:
    """Encoder for serializing spreadsheet data into compact format."""
    
    def __init__(self, max_tokens: int = 4000):
        """Initialize encoder with maximum token limit."""
        self.max_tokens = max_tokens
    
    def encode_sheet(self, sheet_matrix: SheetMatrix) -> str:
        """Encode sheet matrix to Markdown-like format."""
        # Generate header with metadata
        header = self._generate_header(sheet_matrix)
        
        # Generate inverted index of values
        value_index = self._generate_value_index(sheet_matrix)
        
        # Generate format index
        format_index = self._generate_format_index(sheet_matrix)
        
        # Generate merged cells section
        merged_section = self._generate_merged_section(sheet_matrix)
        
        # Combine all sections
        encoded = f"{header}\n\n{value_index}\n\n{format_index}"
        if merged_section:
            encoded += f"\n\n{merged_section}"
            
        # Check token count
        token_count = estimate_tokens(encoded)
        if token_count > self.max_tokens:
            raise ValueError(f"Encoded sheet exceeds token limit: {token_count} > {self.max_tokens}")
            
        return encoded
    
    def _generate_header(self, sheet_matrix: SheetMatrix) -> str:
        """Generate header with sheet metadata."""
        return (
            f"# Sheet: {sheet_matrix.sheet_name}\n"
            f"Dimensions: {sheet_matrix.max_row}x{sheet_matrix.max_col}"
        )
    
    def _generate_value_index(self, sheet_matrix: SheetMatrix) -> str:
        """Generate inverted index of cell values."""
        value_to_cells = defaultdict(list)
        empty_cells = []
        
        # Group cells by value
        for row in range(1, sheet_matrix.max_row + 1):
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell is None or cell.value is None:
                    empty_cells.append(get_cell_address(row, col))
                    continue
                
                # Handle different data types
                if cell.data_type in (DataType.INT_NUM, DataType.FLOAT_NUM):
                    key = f"NUM:{cell.value}"
                elif cell.data_type == DataType.TEXT:
                    key = f"TEXT:{cell.value}"
                else:
                    key = f"{cell.data_type.name}:{cell.value}"
                    
                value_to_cells[key].append(get_cell_address(row, col))
        
        # Convert to serializable format
        sections = []
        
        # Add non-empty cells
        for value_key, cells in value_to_cells.items():
            data_type, value = value_key.split(':', 1)
            sections.append(f"## {data_type}\n{value}\nCells: {', '.join(cells)}")
        
        # Add empty cells section if there are any
        if empty_cells:
            sections.append(f"## EMPTY\nCells: {', '.join(empty_cells)}")
        
        return "\n\n".join(sections)
    
    def _generate_format_index(self, sheet_matrix: SheetMatrix) -> str:
        """Generate index of cell formats."""
        format_to_cells = defaultdict(list)
        
        # Group cells by format
        for row in range(1, sheet_matrix.max_row + 1):
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.format.formats:  # Only include cells with formatting
                    format_key = str(cell.format)
                    format_to_cells[format_key].append(get_cell_address(row, col))
        
        # Convert to serializable format
        if not format_to_cells:
            return "## Formats\nNo special formatting"
            
        sections = ["## Formats"]
        
        for format_str, cells in format_to_cells.items():
            sections.append(f"### {format_str}\nCells: {', '.join(cells)}")
        
        return "\n\n".join(sections)
    
    def _generate_merged_section(self, sheet_matrix: SheetMatrix) -> str:
        """Generate section for merged cells."""
        if not sheet_matrix.merged_cells:
            return ""
            
        merged_ranges = []
        for min_row, min_col, max_row, max_col in sheet_matrix.merged_cells:
            start_cell = get_cell_address(min_row, min_col)
            end_cell = get_cell_address(max_row, max_col)
            merged_ranges.append(f"{start_cell}:{end_cell}")
        
        return f"## Merged Cells\n{', '.join(merged_ranges)}" 