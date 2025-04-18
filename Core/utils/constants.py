"""Constants for the Spreadsheet Processing Framework."""

from enum import Enum, auto

class DataType(Enum):
    """Enumeration of data types for spreadsheet cells."""
    EMPTY = auto()
    INT_NUM = auto()
    FLOAT_NUM = auto()
    PERCENTAGE = auto()
    DATE = auto()
    TIME = auto()
    DATETIME = auto()
    EMAIL = auto()
    PHONE = auto()
    URL = auto()
    TEXT = auto()
    FORMULA = auto()
    CURRENCY = auto()
    BOOLEAN = auto()
    ERROR = auto()
    
class FormatType(Enum):
    """Enumeration of format types for spreadsheet cells."""
    NORMAL = auto()
    BOLD = auto()
    ITALIC = auto()
    UNDERLINE = auto()
    STRIKETHROUGH = auto()
    SUPERSCRIPT = auto()
    SUBSCRIPT = auto()
    CENTER_ALIGNED = auto()
    LEFT_ALIGNED = auto()
    RIGHT_ALIGNED = auto()
    HIGHLIGHTED = auto()
    COLORED = auto()
    BORDERED = auto()
    MERGED = auto() 