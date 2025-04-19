"""Compressor module for reducing spreadsheet size while preserving essential information."""

from typing import Dict, List, Tuple, Set, Optional, Any
import numpy as np
import json
import os
from pathlib import Path
from collections import defaultdict

from Core.parser import SheetMatrix, Cell
from Core.encoder import SpreadsheetEncoder
from Core.utils.constants import DataType, FormatType
from Core.utils.helpers import get_cell_address, estimate_tokens

class CompressionStrategy:
    """Base class for compression strategies."""
    
    def apply(self, sheet_matrix: SheetMatrix) -> SheetMatrix:
        """Apply the compression strategy to the sheet matrix."""
        raise NotImplementedError("Subclasses must implement apply()")
    
    def get_name(self) -> str:
        """Get the name of the compression strategy."""
        return self.__class__.__name__

class StructuralAnchorStrategy(CompressionStrategy):
    """Strategy that extracts content based on structural anchors."""
    
    def __init__(self, anchor_proximity: int = 4):
        """Initialize with anchor proximity."""
        self.anchor_proximity = anchor_proximity
    
    def apply(self, sheet_matrix: SheetMatrix) -> SheetMatrix:
        """Apply structural anchor-based extraction."""
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
        """Find anchor rows based on content heterogeneity."""
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
        """Find anchor columns based on content heterogeneity."""
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

class SheetCompressor:
    """Main compressor class implementing various compression strategies."""
    
    def __init__(self, use_structural_anchors: bool = True, 
                 use_inverted_index: bool = True, 
                 use_format_aggregation: bool = True,
                 anchor_proximity: int = 4,
                 max_tokens: int = 4000):
        """Initialize compressor with selected strategies."""
        self.strategies = []
        self.max_tokens = max_tokens
        
        # Add selected strategies
        if use_structural_anchors:
            self.strategies.append(StructuralAnchorStrategy(anchor_proximity))
            
        # Configure encoder options
        self.use_inverted_index = use_inverted_index
        self.use_format_aggregation = use_format_aggregation
        
    def compress(self, sheet_matrix: SheetMatrix) -> SheetMatrix:
        """Apply all enabled compression strategies to the sheet matrix."""
        compressed_matrix = sheet_matrix
        
        for strategy in self.strategies:
            compressed_matrix = strategy.apply(compressed_matrix)
            
        return compressed_matrix
    
    def compress_and_save(self, sheet_matrix: SheetMatrix, output_dir: str, filename_prefix: str = None) -> Dict:
        """Compress the sheet matrix, encode it, and save to a file.
        
        Args:
            sheet_matrix: The sheet matrix to compress
            output_dir: Directory to save the output file
            filename_prefix: Optional prefix for the output filename
        
        Returns:
            Dict with compression statistics
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Determine filename
        if filename_prefix is None:
            filename_prefix = sheet_matrix.sheet_name.replace(" ", "_")
        
        # Get statistics before compression
        original_cell_count = sum(1 for row in range(1, sheet_matrix.max_row + 1) 
                                for col in range(1, sheet_matrix.max_col + 1) 
                                if sheet_matrix.get_cell(row, col) is not None and 
                                   sheet_matrix.get_cell(row, col).value is not None)
        
        # Apply compression
        compressed_matrix = self.compress(sheet_matrix)
        
        # Get statistics after compression
        compressed_cell_count = sum(1 for row in range(1, compressed_matrix.max_row + 1) 
                                  for col in range(1, compressed_matrix.max_col + 1) 
                                  if compressed_matrix.get_cell(row, col) is not None and 
                                     compressed_matrix.get_cell(row, col).value is not None)
        
        # Encode the compressed sheet
        encoder = SpreadsheetEncoder(max_tokens=self.max_tokens)
        encoder.use_inverted_index = self.use_inverted_index
        encoder.use_format_aggregation = self.use_format_aggregation
        
        try:
            encoded_compressed = encoder.encode_sheet(compressed_matrix)
            compressed_tokens = estimate_tokens(encoded_compressed)
            
            # Skip saving intermediate files - only save final combined output
            output_filename = os.path.join(output_dir, f"{filename_prefix}_combined_all_modules.json")
            
            with open(output_filename, 'w') as f:
                f.write(encoded_compressed)
            
            # Compile statistics
            stats = {
                "sheet_name": sheet_matrix.sheet_name,
                "original_dimensions": f"{sheet_matrix.max_row}x{sheet_matrix.max_col}",
                "original_cell_count": original_cell_count,
                "compressed_cell_count": compressed_cell_count,
                "compression_ratio": compressed_cell_count / original_cell_count,
                "token_count": compressed_tokens,
                "output_file": output_filename,
                "compression_strategies": [strategy.get_name() for strategy in self.strategies],
                "encoding_options": {
                    "inverted_index": self.use_inverted_index,
                    "format_aggregation": self.use_format_aggregation
                }
            }
            
            return stats
            
        except ValueError as e:
            error_msg = f"Error encoding compressed sheet: {str(e)}"
            print(error_msg)
            return {"error": error_msg}

def compress_spreadsheet_to_file(sheet_matrix: SheetMatrix, output_dir: str, filename_prefix: str = None,
                              use_structural_anchors: bool = True, use_inverted_index: bool = True,
                              use_format_aggregation: bool = True, anchor_proximity: int = 4,
                              max_tokens: int = 4000) -> Dict:
    """Convenience function to compress a spreadsheet and save it to a file.
    
    Args:
        sheet_matrix: The sheet matrix to compress
        output_dir: Directory to save the output file
        filename_prefix: Optional prefix for the output filename
        use_structural_anchors: Whether to use structural anchor-based extraction
        use_inverted_index: Whether to use inverted-index translation
        use_format_aggregation: Whether to use format aggregation
        anchor_proximity: Proximity range for structural anchors
        max_tokens: Maximum token limit
    
    Returns:
        Dict with compression statistics
    """
    compressor = SheetCompressor(
        use_structural_anchors=use_structural_anchors,
        use_inverted_index=use_inverted_index,
        use_format_aggregation=use_format_aggregation,
        anchor_proximity=anchor_proximity,
        max_tokens=max_tokens
    )
    
    return compressor.compress_and_save(sheet_matrix, output_dir, filename_prefix)

def extract_metadata(sheet_matrix: SheetMatrix) -> Dict:
    """Extract metadata from sheet matrix.
    
    Args:
        sheet_matrix: The sheet matrix to analyze
        
    Returns:
        Dictionary containing metadata
    """
    metadata = {
        "dimensions": f"{sheet_matrix.max_row}x{sheet_matrix.max_col}",
        "headers": [],
        "data_types": set(),
        "date_ranges": {},
        "key_metrics": []
    }
    
    # Extract headers (first row)
    if sheet_matrix.max_row > 0:
        for col in range(1, sheet_matrix.max_col + 1):
            cell = sheet_matrix.get_cell(1, col)
            if cell and cell.value:
                metadata["headers"].append(str(cell.value))
    
    # Analyze data types and content
    for row in range(1, sheet_matrix.max_row + 1):
        for col in range(1, sheet_matrix.max_col + 1):
            cell = sheet_matrix.get_cell(row, col)
            if cell and cell.value:
                # Track data types
                metadata["data_types"].add(cell.data_type.name)
                
                # Check for dates
                if cell.data_type.name in ["DATE", "DATETIME"]:
                    if "date_ranges" not in metadata:
                        metadata["date_ranges"] = {"min": None, "max": None}
                    if not metadata["date_ranges"]["min"] or cell.value < metadata["date_ranges"]["min"]:
                        metadata["date_ranges"]["min"] = cell.value
                    if not metadata["date_ranges"]["max"] or cell.value > metadata["date_ranges"]["max"]:
                        metadata["date_ranges"]["max"] = cell.value
                
                # Check for key metrics
                if any(keyword in str(cell.value).lower() for keyword in ["revenue", "income", "expense", "profit", "loss"]):
                    metadata["key_metrics"].append(str(cell.value))
    
    # Convert sets to lists for JSON serialization
    metadata["data_types"] = list(metadata["data_types"])
    
    return metadata

def compress_with_best_method(sheet_matrix: SheetMatrix, output_dir: str, filename_prefix: str = None,
                            max_tokens: int = 4000) -> Dict:
    """
    Automatically select and apply the best compression method for the given sheet matrix.
    Optimized for runtime performance while maintaining output quality.
    
    Args:
        sheet_matrix: The sheet matrix to compress
        output_dir: Directory to save the compressed output
        filename_prefix: Optional prefix for the output filename
        max_tokens: Maximum number of tokens allowed in the output
        
    Returns:
        Dict containing compression results and metadata
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Extract metadata
    metadata = extract_metadata(sheet_matrix)
    
    # Initialize metadata index
    metadata_index = MetadataIndex(output_dir)
    if filename_prefix:
        metadata_index.update_index(filename_prefix, metadata)
    
    # Initialize encoder with token limit
    encoder = SpreadsheetEncoder(max_tokens=max_tokens)
    
    # Quick size check - if sheet is very large, use Module 3 directly
    estimated_size = sheet_matrix.max_row * sheet_matrix.max_col * 2
    if estimated_size > 100000:  # Large sheet threshold
        try:
            encoder.use_structural_anchors = False
            encoder.use_inverted_index = False
            encoder.use_format_aggregation = True
            encoded = encoder.encode_sheet(sheet_matrix)
            
            # Generate output filename
            if filename_prefix is None:
                filename_prefix = sheet_matrix.sheet_name
            output_file = os.path.join(output_dir, f"{filename_prefix}_compressed.json")
            
            # Save the output
            with open(output_file, 'w') as f:
                f.write(encoded)
            
            return {
                "method": "module3",
                "compression_ratio": 1.0,  # Not calculated for large sheets
                "output_file": output_file,
                "tokens": estimate_tokens(encoded),
                "metadata": metadata
            }
        except Exception as e:
            print(f"Error with Module 3 for large sheet: {str(e)}")
            raise
    
    # For smaller sheets, try all methods in order of expected effectiveness
    methods = [
        ("combined", {"use_structural_anchors": True, "use_inverted_index": True, "use_format_aggregation": True}),
        ("module3", {"use_structural_anchors": False, "use_inverted_index": False, "use_format_aggregation": True}),
        ("module2", {"use_structural_anchors": False, "use_inverted_index": True, "use_format_aggregation": False}),
        ("module1", {"use_structural_anchors": True, "use_inverted_index": False, "use_format_aggregation": False})
    ]
    
    # Cache the original encoding to avoid recalculating
    original_encoded = encoder.encode_sheet(sheet_matrix)
    original_tokens = estimate_tokens(original_encoded)
    
    best_ratio = float('inf')
    best_method = None
    best_encoded = None
    
    # Try each method and measure compression ratio
    for method_name, params in methods:
        try:
            # Configure encoder
            encoder.use_structural_anchors = params["use_structural_anchors"]
            encoder.use_inverted_index = params["use_inverted_index"]
            encoder.use_format_aggregation = params["use_format_aggregation"]
            
            # Apply compression
            if params["use_structural_anchors"]:
                compressor = SheetCompressor(**params)
                compressed_matrix = compressor.compress(sheet_matrix)
                encoded = encoder.encode_sheet(compressed_matrix)
            else:
                encoded = encoder.encode_sheet(sheet_matrix)
            
            # Calculate compression ratio
            compressed_tokens = estimate_tokens(encoded)
            ratio = compressed_tokens / original_tokens if original_tokens > 0 else 0
            
            # If we achieve very good compression, use it immediately
            if ratio < 0.3:  # 70% compression or better
                best_ratio = ratio
                best_method = method_name
                best_encoded = encoded
                break
                
            # Update best method if this one is better
            if ratio < best_ratio:
                best_ratio = ratio
                best_method = method_name
                best_encoded = encoded
                
        except Exception as e:
            print(f"Error with {method_name}: {str(e)}")
            continue
    
    if best_encoded is None:
        raise ValueError("No compression method succeeded")
    
    # Generate output filename
    if filename_prefix is None:
        filename_prefix = sheet_matrix.sheet_name
    output_file = os.path.join(output_dir, f"{filename_prefix}_compressed.json")
    
    # Save the best compressed output
    with open(output_file, 'w') as f:
        f.write(best_encoded)
    
    # Add metadata to the return value
    result = {
        "method": best_method,
        "compression_ratio": best_ratio,
        "output_file": output_file,
        "tokens": estimate_tokens(best_encoded),
        "metadata": metadata
    }
    
    return result 