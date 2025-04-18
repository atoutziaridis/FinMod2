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
        """
        Find anchor rows based on content heterogeneity.
        
        Detects rows that:
        1. Have high diversity in cell values or formats (headers, section breaks)
        2. Contain format changes (borders, colors, font styles)
        3. Mark natural table boundaries
        """
        anchor_rows = set()
        row_scores = {}
        
        # Calculate row heterogeneity scores
        for row in range(1, sheet_matrix.max_row + 1):
            # Track value changes
            value_changes = 0
            unique_values = set()
            format_changes = 0
            has_borders = False
            has_colors = False
            has_emphasis = False
            
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    unique_values.add(str(cell.value))
                    
                    # Check format features
                    if cell.format:
                        if cell.format.borders:
                            has_borders = True
                        if cell.format.bg_color or cell.format.font_color:
                            has_colors = True
                        if FormatType.BOLD in cell.format.formats or FormatType.ITALIC in cell.format.formats:
                            has_emphasis = True
            
            # Calculate heterogeneity score
            value_changes = len(unique_values)
            
            # Weighted score following the framework's guidelines
            row_score = (
                value_changes * 2.0 +                # Value diversity weight
                (1.5 if has_borders else 0) +        # Border weight
                (1.0 if has_colors else 0) +         # Color weight
                (1.0 if has_emphasis else 0)         # Font emphasis weight
            )
            
            row_scores[row] = row_score
            
            # Consider row as anchor if score exceeds threshold
            # Threshold determined based on empirical analysis
            if row_score >= 3.0:
                anchor_rows.add(row)
        
        # Always include first and last row as anchors
        if sheet_matrix.max_row > 0:
            anchor_rows.add(1)
            anchor_rows.add(sheet_matrix.max_row)
            
        # If too few anchors detected, lower threshold and retry
        if len(anchor_rows) < 3 and sheet_matrix.max_row > 5:
            secondary_threshold = 2.0
            for row, score in row_scores.items():
                if score >= secondary_threshold:
                    anchor_rows.add(row)
                    
        return anchor_rows
    
    def _find_anchor_cols(self, sheet_matrix: SheetMatrix) -> Set[int]:
        """
        Find anchor columns based on content heterogeneity.
        
        Detects columns that:
        1. Have high diversity in cell values or formats
        2. Contain format changes (borders, colors, font styles)
        3. Mark natural table boundaries
        """
        anchor_cols = set()
        col_scores = {}
        
        # Calculate column heterogeneity scores
        for col in range(1, sheet_matrix.max_col + 1):
            # Track value changes
            value_changes = 0
            unique_values = set()
            has_borders = False
            has_colors = False
            has_emphasis = False
            has_nonumeric = 0
            total_cells = 0
            
            for row in range(1, sheet_matrix.max_row + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    total_cells += 1
                    unique_values.add(str(cell.value))
                    
                    # Check for non-numeric values (headers often contain text)
                    if cell.data_type not in [DataType.INT_NUM, DataType.FLOAT_NUM]:
                        has_nonumeric += 1
                    
                    # Check format features
                    if cell.format:
                        if cell.format.borders:
                            has_borders = True
                        if cell.format.bg_color or cell.format.font_color:
                            has_colors = True
                        if FormatType.BOLD in cell.format.formats or FormatType.ITALIC in cell.format.formats:
                            has_emphasis = True
            
            # Calculate heterogeneity score
            value_changes = len(unique_values)
            
            # Non-numeric ratio (for identifying label columns)
            nonumeric_ratio = has_nonumeric / total_cells if total_cells > 0 else 0
            
            # Weighted score following the framework's guidelines
            col_score = (
                value_changes * 1.5 +                     # Value diversity weight
                (nonumeric_ratio * 3.0) +                 # Non-numeric weight (higher for label columns)
                (1.5 if has_borders else 0) +             # Border weight
                (1.0 if has_colors else 0) +              # Color weight
                (1.0 if has_emphasis else 0)              # Font emphasis weight
            )
            
            col_scores[col] = col_score
            
            # Consider column as anchor if score exceeds threshold
            if col_score >= 3.0:
                anchor_cols.add(col)
        
        # Always include first and last column as anchors
        if sheet_matrix.max_col > 0:
            anchor_cols.add(1)
            anchor_cols.add(sheet_matrix.max_col)
            
        # If too few anchors detected, lower threshold and retry
        if len(anchor_cols) < 3 and sheet_matrix.max_col > 5:
            secondary_threshold = 2.0
            for col, score in col_scores.items():
                if score >= secondary_threshold:
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