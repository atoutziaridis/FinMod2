"""Encoder module for serializing spreadsheet data into token-efficient format."""

from typing import Dict, List, Tuple, Any, Optional, Set, Union
from collections import defaultdict
import json
import re
import math

from Core.parser import SheetMatrix, Cell
from Core.utils.constants import DataType, FormatType
from Core.utils.helpers import estimate_tokens, get_cell_address, detect_number_format_string

class SpreadsheetEncoder:
    """Encoder for serializing spreadsheet data into compact format."""
    
    def __init__(self, max_tokens: int = 4000):
        """Initialize encoder with maximum token limit."""
        self.max_tokens = max_tokens
        self.chunk_size = 100  # Number of rows per chunk
        self.use_inverted_index = True  # Enable/disable inverted index
        self.use_format_aggregation = True  # Enable/disable format aggregation
    
    def encode_sheet(self, sheet_matrix: SheetMatrix) -> str:
        """Encode sheet matrix to Markdown-like format."""
        # Generate header with metadata
        header = self._generate_header(sheet_matrix)
        
        # Try standard encoding first
        try:
            if self.use_inverted_index:
                # Use the inverted index approach
                text_section = self._generate_inverted_index_text(sheet_matrix)
            else:
                # Use the original approach
                text_section = self._generate_text_section(sheet_matrix)
                
            if self.use_format_aggregation:
                # Use format aggregation
                format_section = self._generate_aggregated_format_section(sheet_matrix)
            else:
                # Use the original approach
                format_section = self._generate_format_section(sheet_matrix)
            
            # Combine all sections
            encoded = f"{header}\n\nText Input:\n{text_section}\n\nFormat Input:\n{format_section}"
            
            # Check token count
            token_count = estimate_tokens(encoded)
            if token_count <= self.max_tokens:
                return encoded
            
            # If we're here, encoding exceeded max_tokens
            raise ValueError(f"Encoded sheet exceeds token limit: {token_count} > {self.max_tokens}")
            
        except ValueError:
            # Standard encoding is too large, fall back to chunked encoding
            return self._chunk_encode_sheet(sheet_matrix, header)
    
    def _chunk_encode_sheet(self, sheet_matrix: SheetMatrix, header: str) -> str:
        """Encode large sheets by breaking them into manageable chunks."""
        print(f"Sheet too large for standard encoding, using chunked approach...")
        
        # Try to detect header rows
        header_rows = self._detect_header_rows(sheet_matrix)
        
        # Determine chunk size based on sheet dimensions
        max_body_rows = sheet_matrix.max_row - len(header_rows)
        rows_per_chunk = min(max(20, math.floor(self.max_tokens / 100)), 200)
        num_chunks = math.ceil(max_body_rows / rows_per_chunk)
        
        print(f"Splitting into {num_chunks} chunks with {len(header_rows)} header rows")
        
        chunks = []
        # Always include header rows in each chunk for context
        header_indices = sorted(list(header_rows))
        
        # Create sub-matrices for each chunk
        start_row = max(header_indices[-1] + 1, 1) if header_indices else 1
        for chunk_idx in range(num_chunks):
            chunk_start = start_row + (chunk_idx * rows_per_chunk)
            chunk_end = min(chunk_start + rows_per_chunk - 1, sheet_matrix.max_row)
            
            # Skip empty chunks
            if chunk_start > sheet_matrix.max_row:
                continue
                
            # Create rows to include (headers + chunk rows)
            rows_to_include = set(header_indices).union(range(chunk_start, chunk_end + 1))
            
            # Create a sub-matrix for this chunk
            chunk_matrix = self._create_sub_matrix(sheet_matrix, rows_to_include)
            
            # Encode the chunk
            if self.use_inverted_index:
                chunk_text = self._generate_inverted_index_text(chunk_matrix)
            else:
                chunk_text = self._generate_text_section(chunk_matrix)
                
            if self.use_format_aggregation:
                chunk_format = self._generate_aggregated_format_section(chunk_matrix)
            else:
                chunk_format = self._generate_format_section(chunk_matrix)
            
            chunk_encoded = f"Chunk {chunk_idx+1}/{num_chunks} (rows {chunk_start}-{chunk_end}):\n\nText Input:\n{chunk_text}\n\nFormat Input:\n{chunk_format}"
            chunks.append(chunk_encoded)
        
        # Combine chunks with metadata
        return f"{header}\n\n" + "\n\n---\n\n".join(chunks)
    
    def _create_sub_matrix(self, sheet_matrix: SheetMatrix, rows_to_include: Set[int]) -> SheetMatrix:
        """Create a sub-matrix with only the specified rows."""
        # Create a new matrix with only the specified rows
        new_matrix = []
        for row in range(1, sheet_matrix.max_row + 1):
            if row in rows_to_include:
                new_row = []
                for col in range(1, sheet_matrix.max_col + 1):
                    new_row.append(sheet_matrix.get_cell(row, col))
                new_matrix.append(new_row)
            else:
                new_matrix.append([None] * sheet_matrix.max_col)
        
        # Create new SheetMatrix
        return SheetMatrix(
            matrix=new_matrix,
            merged_cells=sheet_matrix.merged_cells,  # Keep all merged cells for simplicity
            sheet_name=sheet_matrix.sheet_name,
            max_row=sheet_matrix.max_row,
            max_col=sheet_matrix.max_col
        )
    
    def _detect_header_rows(self, sheet_matrix: SheetMatrix) -> Set[int]:
        """Detect likely header rows in the sheet."""
        header_rows = set()
        
        # First row is often a header
        if sheet_matrix.max_row > 0:
            header_rows.add(1)
        
        # Look for rows with formatting indicators of headers
        for row in range(1, min(10, sheet_matrix.max_row) + 1):  # Check first 10 rows
            has_bold = False
            has_borders = False
            text_cells = 0
            total_cells = 0
            
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    total_cells += 1
                    
                    # Check for text cells (headers are usually text)
                    if cell.data_type == DataType.TEXT:
                        text_cells += 1
                    
                    # Check for formatting
                    if cell.format:
                        if FormatType.BOLD in cell.format.formats:
                            has_bold = True
                        if cell.format.borders:
                            has_borders = True
            
            # If row has bold cells, borders, and mostly text, it's likely a header
            if has_bold or has_borders:
                if total_cells > 0 and text_cells / total_cells >= 0.5:
                    header_rows.add(row)
        
        return header_rows
    
    def _generate_header(self, sheet_matrix: SheetMatrix) -> str:
        """Generate header with sheet metadata."""
        return f"# Sheet: {sheet_matrix.sheet_name}\nDimensions: {sheet_matrix.max_row}x{sheet_matrix.max_col}"
    
    def _generate_text_section(self, sheet_matrix: SheetMatrix) -> str:
        """Generate text section with cell values."""
        rows = []
        
        for row in range(1, sheet_matrix.max_row + 1):
            row_data = []
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    row_data.append(f"{get_cell_address(row, col)},{cell.value}")
                else:
                    row_data.append(f"{get_cell_address(row, col)},")
            
            # Join cells in the row with | separator
            row_str = "|".join(row_data)
            rows.append(row_str)
        
        return "\n".join(rows)
    
    def _generate_inverted_index_text(self, sheet_matrix: SheetMatrix) -> str:
        """Generate inverted index of cell values to addresses."""
        # Step 1: Build the inverted index, optimized to group by value types
        value_to_cells = defaultdict(list)
        empty_count = 0
        
        for row in range(1, sheet_matrix.max_row + 1):
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    # Prefix values with data type for better grouping
                    if cell.data_type in [DataType.INT_NUM, DataType.FLOAT_NUM]:
                        # For numbers, use shorter representation to save tokens
                        prefix = "N"  # N for number
                    else:
                        prefix = cell.data_type.name[0]  # First letter of data type
                    
                    # Convert value to string representation
                    value_str = str(cell.value)
                    # Use key with type prefix for consistent grouping
                    key = f"{prefix}:{value_str}"
                    value_to_cells[key].append(get_cell_address(row, col))
                else:
                    empty_count += 1
        
        # Step 2: Compress address ranges
        compressed_index = {}
        for key, addresses in value_to_cells.items():
            # Skip the type prefix for the output
            output_key = key.split(":", 1)[1]
            compressed_ranges = self._compress_address_ranges(addresses)
            compressed_index[output_key] = compressed_ranges
        
        # Add empty cell count instead of listing all empty cells
        if empty_count > 0:
            compressed_index["__empty_cells"] = empty_count
        
        # Step 3: Generate JSON representation, minimize whitespace
        return json.dumps(compressed_index, separators=(',', ':'))
    
    def _compress_address_ranges(self, addresses: List[str]) -> List[str]:
        """Compress contiguous cell addresses into ranges."""
        if not addresses:
            return []
        
        # Sort addresses to help find contiguous cells
        addresses.sort()
        
        # First pass: Find row-wise ranges
        ranges = []
        current_range = [addresses[0]]
        current_row = ''.join(c for c in addresses[0] if c.isdigit())
        current_col_base = ''.join(c for c in addresses[0] if c.isalpha())
        current_col_num = self._column_letter_to_number(current_col_base)
        
        for i in range(1, len(addresses)):
            addr = addresses[i]
            row = ''.join(c for c in addr if c.isdigit())
            col_base = ''.join(c for c in addr if c.isalpha())
            col_num = self._column_letter_to_number(col_base)
            
            # Check if it's the next column in the same row
            if row == current_row and col_num == current_col_num + 1:
                current_range.append(addr)
                current_col_num = col_num
            else:
                # This cell is not contiguous, end the current range
                if len(current_range) >= 3:  # Only compress if 3+ cells
                    ranges.append(f"{current_range[0]}:{current_range[-1]}")
                else:
                    ranges.extend(current_range)
                
                # Start a new range
                current_range = [addr]
                current_row = row
                current_col_base = col_base
                current_col_num = col_num
        
        # Handle the last range
        if len(current_range) >= 3:
            ranges.append(f"{current_range[0]}:{current_range[-1]}")
        else:
            ranges.extend(current_range)
        
        # Second pass: Find column-wise ranges (could be implemented similarly)
        
        return ranges
    
    def _column_letter_to_number(self, column_letter: str) -> int:
        """Convert Excel column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)"""
        result = 0
        for c in column_letter:
            result = result * 26 + (ord(c.upper()) - ord('A') + 1)
        return result
    
    def _generate_format_section(self, sheet_matrix: SheetMatrix) -> str:
        """Generate format section with cell formatting."""
        rows = []
        
        for row in range(1, sheet_matrix.max_row + 1):
            row_data = []
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.format:
                    # Convert format types to string representation
                    format_str = ",".join(fmt.name for fmt in cell.format.formats)
                    if cell.format.bg_color:
                        format_str += f",Fill Color"
                    if cell.format.font_color:
                        format_str += f",Font Color"
                    if cell.format.borders:
                        format_str += "," + ",".join(f"{border} Border" for border in cell.format.borders)
                    row_data.append(f"{get_cell_address(row, col)},{format_str}")
                else:
                    row_data.append(f"{get_cell_address(row, col)},")
            
            # Join cells in the row with | separator
            row_str = "|".join(row_data)
            rows.append(row_str)
        
        return "\n".join(rows)
    
    def _generate_aggregated_format_section(self, sheet_matrix: SheetMatrix) -> str:
        """Generate aggregated format section, grouping cells by format and data type."""
        # 1. Group by format + number format string
        format_groups = defaultdict(list)
        data_type_groups = defaultdict(list)
        
        for row in range(1, sheet_matrix.max_row + 1):
            for col in range(1, sheet_matrix.max_col + 1):
                cell = sheet_matrix.get_cell(row, col)
                if cell and cell.value is not None:
                    cell_address = get_cell_address(row, col)
                    
                    # Group by data type and format string
                    if cell.data_type in [DataType.INT_NUM, DataType.FLOAT_NUM, DataType.PERCENTAGE, DataType.CURRENCY]:
                        # Get the number format string
                        format_string = detect_number_format_string(cell.value)
                        if format_string:
                            key = f"NFS:{format_string}"
                            data_type_groups[key].append(cell_address)
                    
                    # Group by cell format
                    if cell.format:
                        format_str = str(cell.format)
                        if format_str:  # Skip empty formats
                            format_groups[format_str].append(cell_address)
        
        # 2. Detect rectangular blocks for better compression
        aggregated_formats = {}
        
        # Process format groups
        for format_str, addresses in format_groups.items():
            if len(addresses) > 5:  # Only try to find rectangles for larger groups
                rectangles = self._find_rectangles(addresses)
                if rectangles:
                    compressed = rectangles
                else:
                    compressed = self._compress_address_ranges(addresses)
            else:
                compressed = self._compress_address_ranges(addresses)
            
            aggregated_formats[f"FMT:{format_str}"] = compressed
        
        # Process data type groups
        for data_type, addresses in data_type_groups.items():
            if len(addresses) > 5:
                rectangles = self._find_rectangles(addresses)
                if rectangles:
                    compressed = rectangles
                else:
                    compressed = self._compress_address_ranges(addresses)
            else:
                compressed = self._compress_address_ranges(addresses)
            
            aggregated_formats[data_type] = compressed
        
        return json.dumps(aggregated_formats, separators=(',', ':'))
    
    def _find_rectangles(self, addresses: List[str]) -> List[str]:
        """Find rectangular blocks of cells from a list of addresses."""
        if not addresses or len(addresses) < 4:  # Need at least 4 cells for a rectangle
            return []
        
        # Parse addresses to (row, col) coordinates
        coordinates = []
        for addr in addresses:
            col_str = ''.join(c for c in addr if c.isalpha())
            row = int(''.join(c for c in addr if c.isdigit()))
            col = self._column_letter_to_number(col_str)
            coordinates.append((row, col, addr))
        
        # Sort by row, then by column
        coordinates.sort()
        
        # Try to find rectangles
        rectangles = []
        remaining = set(addr for _, _, addr in coordinates)
        
        # For each potential top-left corner
        for start_idx, (start_row, start_col, start_addr) in enumerate(coordinates):
            if start_addr not in remaining:
                continue
                
            # Find the maximum width (all cells in the same row)
            max_width = 1
            for i in range(start_idx + 1, len(coordinates)):
                row, col, addr = coordinates[i]
                if row == start_row and col == start_col + max_width and addr in remaining:
                    max_width += 1
                elif row > start_row:
                    break
            
            if max_width < 2:  # Need at least 2 columns for a rectangle
                continue
                
            # Find the maximum height (all cells in the rectangle)
            max_height = 1
            is_rectangle = True
            
            for height in range(1, 100):  # Reasonable limit
                current_row = start_row + height
                # Check if all cells in this row exist
                for width in range(max_width):
                    current_col = start_col + width
                    # Find the address for this coordinate
                    current_addr = None
                    for r, c, addr in coordinates:
                        if r == current_row and c == current_col:
                            current_addr = addr
                            break
                    
                    if current_addr is None or current_addr not in remaining:
                        is_rectangle = False
                        break
                
                if not is_rectangle:
                    break
                    
                max_height = height + 1
            
            if max_height < 2:  # Need at least 2 rows for a rectangle
                continue
                
            # We found a rectangle
            end_row = start_row + max_height - 1
            end_col = start_col + max_width - 1
            end_addr = get_cell_address(end_row, end_col)
            
            rectangle = f"{start_addr}:{end_addr}"
            rectangles.append(rectangle)
            
            # Remove these addresses from consideration
            for r in range(start_row, start_row + max_height):
                for c in range(start_col, start_col + max_width):
                    for coord_row, coord_col, addr in coordinates:
                        if coord_row == r and coord_col == c:
                            remaining.discard(addr)
        
        # Add any remaining addresses as individual cells or ranges
        if remaining:
            remaining_list = list(remaining)
            remaining_compressed = self._compress_address_ranges(remaining_list)
            rectangles.extend(remaining_compressed)
        
        return rectangles
    
    def _calculate_dynamic_token_limit(self, sheet_matrix: SheetMatrix) -> int:
        """Calculate dynamic token limit based on sheet size."""
        base_tokens = 1000  # Base tokens for header and metadata
        cell_tokens = sheet_matrix.max_row * sheet_matrix.max_col * 2  # Approximate tokens per cell
        return min(self.max_tokens, base_tokens + cell_tokens) 