"""Core module for the Spreadsheet Processing Framework."""

from Core.parser import SpreadsheetParser
from Core.compressor import SheetCompressor, compress_spreadsheet_to_file
from Core.encoder import SpreadsheetEncoder

# Make the most important functions and classes available at the top level
__all__ = [
    'SpreadsheetParser', 
    'SheetCompressor', 
    'SpreadsheetEncoder',
    'compress_spreadsheet_to_file'
] 