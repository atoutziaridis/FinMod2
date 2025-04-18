"""Compressor module for reducing spreadsheet size while preserving essential information."""

from typing import Dict, List, Tuple, Set, Optional, Any
import numpy as np
from collections import defaultdict

from Core.parser import SheetMatrix, Cell
from Core.utils.constants import DataType, FormatType
from Core.utils.helpers import get_cell_address

class SheetCompressor:
    """Main compressor class implementing various compression strategies."""
    
    def __init__(self, use_structural_anchors: bool = True, 
                 use_inverted_index: bool = True, 
                 use_format_aggregation: bool = True,
                 anchor_proximity: int = 4):
        """Initialize compressor with selected strategies."""
        self.use_structural_anchors = use_structural_anchors
        self.use_inverted_index = use_inverted_index
        self.use_format_aggregation = use_format_aggregation
        self.anchor_proximity = anchor_proximity
        
    def compress(self, sheet_matrix: SheetMatrix) -> SheetMatrix:
        """Apply all enabled compression strategies to the sheet matrix."""
        compressed_matrix = sheet_matrix
        
        # Apply structural anchor-based extraction
        if self.use_structural_anchors:
            compressed_matrix = self._structural_anchor_extraction(compressed_matrix)
            
        # Apply inverted index translation (to be implemented)
        if self.use_inverted_index:
            pass  # Will implement this next
            
        # Apply data-format-aware aggregation (to be implemented)
        if self.use_format_aggregation:
            pass  # Will implement this next
            
        return compressed_matrix
        
    def _structural_anchor_extraction(self, sheet_matrix: SheetMatrix) -> SheetMatrix:
        """
        Module 1: Structural-Anchor-Based Extraction
        
        Detect boundary rows/columns based on cell content variation,
        and retain rows/cols within k of these anchors.
        """
        # Find anchor rows and columns
        anchor_rows = self._find_anchor_rows(sheet_matrix)
        anchor_cols = self._find_anchor_cols(sheet_matrix)
        
        # Expand anchors by proximity
        rows_to_keep = self._expand_anchors(anchor_rows, sheet_matrix.max_row)
        cols_to_keep = self._expand_anchors(anchor_cols, sheet_matrix.max_col)
        
        # Create new reduced matrix
        new_matrix = []
        for row in range(1, sheet_matrix.max_row + 1):
            if row in rows_to_keep:
                new_row = []
                for col in range(1, sheet_matrix.max_col + 1):
                    if col in cols_to_keep:
                        new_row.append(sheet_matrix.get_cell(row, col))
                    else:
                        new_row.append(None)
                new_matrix.append(new_row)
            else:
                new_matrix.append([None] * sheet_matrix.max_col)
        
        # Update merged cells to reflect the compression
        new_merged_cells = []
        for min_row, min_col, max_row, max_col in sheet_matrix.merged_cells:
            if (min_row in rows_to_keep or max_row in rows_to_keep) and \
               (min_col in cols_to_keep or max_col in cols_to_keep):
                new_merged_cells.append((min_row, min_col, max_row, max_col))
        
        # Create new SheetMatrix with reduced data
        return SheetMatrix(
            matrix=new_matrix,
            merged_cells=new_merged_cells,
            sheet_name=sheet_matrix.sheet_name,
            max_row=sheet_matrix.max_row,
            max_col=sheet_matrix.max_col
        )
    
    def _find_anchor_rows(self, sheet_matrix: SheetMatrix) -> Set[int]:
        """Find anchor rows based on cell content variation and formatting."""
        anchor_rows = set()
        
        # Track data type changes
        prev_row_types = []
        
        for row in range(1, sheet_matrix.max_row + 1):
            # Check for header/title indicators
            has_bold = False
            has_border = False
            has_highlight = False
            
            # Check for data type variation
            row_types = []
            
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell:
                    # Check for formatting indicators
                    if cell.format.formats:
                        if FormatType.BOLD in cell.format.formats:
                            has_bold = True
                        if FormatType.BORDERED in cell.format.formats:
                            has_border = True
                        if FormatType.HIGHLIGHTED in cell.format.formats:
                            has_highlight = True
                    
                    # Track data type
                    if cell.value is not None:
                        row_types.append(cell.data_type)
            
            # Determine if this is an anchor row
            
            # Check for formatting indicators of headers/titles
            if has_bold or has_border or has_highlight:
                anchor_rows.add(row)
                continue
            
            # Check for data type transitions
            if prev_row_types and row_types:
                # Check if this row marks a transition between data types
                if set(row_types) != set(prev_row_types):
                    anchor_rows.add(row)
                    
            # Store current row's types for next iteration
            prev_row_types = row_types
            
            # Also check for non-empty rows surrounded by empty rows
            prev_empty = row > 1 and all(sheet_matrix.get_cell(row-1, col) is None or 
                                        sheet_matrix.get_cell(row-1, col).value is None 
                                        for col in range(1, sheet_matrix.max_col + 1))
            
            next_empty = row < sheet_matrix.max_row and all(sheet_matrix.get_cell(row+1, col) is None or 
                                                          sheet_matrix.get_cell(row+1, col).value is None 
                                                          for col in range(1, sheet_matrix.max_col + 1))
            
            current_non_empty = any(sheet_matrix.get_cell(row, col) is not None and 
                                   sheet_matrix.get_cell(row, col).value is not None 
                                   for col in range(1, sheet_matrix.max_col + 1))
            
            if current_non_empty and (prev_empty or next_empty):
                anchor_rows.add(row)
        
        return anchor_rows
    
    def _find_anchor_cols(self, sheet_matrix: SheetMatrix) -> Set[int]:
        """Find anchor columns based on cell content variation and formatting."""
        anchor_cols = set()
        
        # Track data type changes
        prev_col_types = []
        
        for col in range(1, sheet_matrix.max_col + 1):
            # Check for header/title indicators
            has_bold = False
            has_border = False
            has_highlight = False
            
            # Check for data type variation
            col_types = []
            
            for row in range(1, sheet_matrix.max_row + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell:
                    # Check for formatting indicators
                    if cell.format.formats:
                        if FormatType.BOLD in cell.format.formats:
                            has_bold = True
                        if FormatType.BORDERED in cell.format.formats:
                            has_border = True
                        if FormatType.HIGHLIGHTED in cell.format.formats:
                            has_highlight = True
                    
                    # Track data type
                    if cell.value is not None:
                        col_types.append(cell.data_type)
            
            # Determine if this is an anchor column
            
            # Check for formatting indicators of headers/titles
            if has_bold or has_border or has_highlight:
                anchor_cols.add(col)
                continue
            
            # Check for data type transitions
            if prev_col_types and col_types:
                # Check if this column marks a transition between data types
                if set(col_types) != set(prev_col_types):
                    anchor_cols.add(col)
                    
            # Store current column's types for next iteration
            prev_col_types = col_types
            
            # Also check for non-empty columns surrounded by empty columns
            prev_empty = col > 1 and all(sheet_matrix.get_cell(row, col-1) is None or 
                                        sheet_matrix.get_cell(row, col-1).value is None 
                                        for row in range(1, sheet_matrix.max_row + 1))
            
            next_empty = col < sheet_matrix.max_col and all(sheet_matrix.get_cell(row, col+1) is None or 
                                                          sheet_matrix.get_cell(row, col+1).value is None 
                                                          for row in range(1, sheet_matrix.max_row + 1))
            
            current_non_empty = any(sheet_matrix.get_cell(row, col) is not None and 
                                   sheet_matrix.get_cell(row, col).value is not None 
                                   for row in range(1, sheet_matrix.max_row + 1))
            
            if current_non_empty and (prev_empty or next_empty):
                anchor_cols.add(col)
        
        return anchor_cols
    
    def _expand_anchors(self, anchors: Set[int], max_index: int) -> Set[int]:
        """Expand anchors by including k cells on either side."""
        expanded = set(anchors)
        
        for anchor in anchors:
            # Add k cells before and after the anchor
            for i in range(1, self.anchor_proximity + 1):
                if anchor - i >= 1:
                    expanded.add(anchor - i)
                if anchor + i <= max_index:
                    expanded.add(anchor + i)
        
        return expanded 