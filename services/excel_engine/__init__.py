"""Excel Engine - XLSX to JSON conversion and back.

This module handles:
1. Parsing XLSX files into an editable JSON structure (ExcelWorkbookJSON)
2. Applying JSON edits back to XLSX files
3. Preserving formatting, merged cells, data validation, images, and structure
"""

from .schemas import (
    # Core
    ExcelWorkbookJSON,
    ExcelSheetJSON,
    ExcelCellJSON,
    MergedCellRange,
    DataValidationRule,
    ExcelImage,
    ExcelComment,
    CellStyle,
    CellBorder,
    CellFill,
    CellFont,
    # Complex elements
    ExcelHyperlink,
    ConditionalFormatting,
    ConditionalFormatRule,
    FormControl,
    FormControlType,
    DefinedName,
    SheetView,
    FreezePane,
    ExcelTable,
    TableColumn,
    Sparkline,
    SparklineGroup,
)
from .parser import xlsx_to_json, parse_sheet
from .writer import apply_json_to_xlsx

__all__ = [
    # Core schemas
    "ExcelWorkbookJSON",
    "ExcelSheetJSON", 
    "ExcelCellJSON",
    "MergedCellRange",
    "DataValidationRule",
    "ExcelImage",
    "ExcelComment",
    "CellStyle",
    "CellBorder",
    "CellFill",
    "CellFont",
    # Complex element schemas
    "ExcelHyperlink",
    "ConditionalFormatting",
    "ConditionalFormatRule",
    "FormControl",
    "FormControlType",
    "DefinedName",
    "SheetView",
    "FreezePane",
    "ExcelTable",
    "TableColumn",
    "Sparkline",
    "SparklineGroup",
    # Functions
    "xlsx_to_json",
    "parse_sheet",
    "apply_json_to_xlsx",
]
