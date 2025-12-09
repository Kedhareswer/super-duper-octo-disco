"""Pydantic schemas for Excel document structure.

These schemas model the editable representation of an XLSX file,
preserving structural fidelity for complex features like:
- Merged cells
- Data validation (dropdowns)
- Cell formatting and styles
- Images and drawings
- Comments/notes
"""

from __future__ import annotations

from enum import Enum
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel, Field


class CellDataType(str, Enum):
    """Excel cell data types."""
    STRING = "s"  # Shared string
    NUMBER = "n"  # Number
    BOOLEAN = "b"  # Boolean
    ERROR = "e"  # Error
    INLINE_STRING = "inlineStr"  # Inline string (not shared)
    FORMULA = "str"  # Formula result as string
    DATE = "d"  # Date (ISO 8601)


class CellFont(BaseModel):
    """Font styling for a cell."""
    name: Optional[str] = None
    size: Optional[float] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: Optional[str] = None  # Hex color e.g. "FF0000"


class CellFill(BaseModel):
    """Fill/background for a cell."""
    pattern_type: Optional[str] = None  # "solid", "none", etc.
    fg_color: Optional[str] = None  # Foreground color (hex)
    bg_color: Optional[str] = None  # Background color (hex)


class CellBorder(BaseModel):
    """Border for a single edge."""
    style: Optional[str] = None  # "thin", "medium", "thick", "dashed", etc.
    color: Optional[str] = None  # Hex color


class CellBorders(BaseModel):
    """All borders for a cell."""
    left: Optional[CellBorder] = None
    right: Optional[CellBorder] = None
    top: Optional[CellBorder] = None
    bottom: Optional[CellBorder] = None
    diagonal: Optional[CellBorder] = None


class CellAlignment(BaseModel):
    """Text alignment in a cell."""
    horizontal: Optional[str] = None  # "left", "center", "right", "justify"
    vertical: Optional[str] = None  # "top", "center", "bottom"
    wrap_text: bool = False
    text_rotation: Optional[int] = None  # 0-180 degrees
    indent: Optional[int] = None


class CellStyle(BaseModel):
    """Complete cell style reference."""
    style_id: Optional[int] = None  # Reference to styles.xml
    font: Optional[CellFont] = None
    fill: Optional[CellFill] = None
    borders: Optional[CellBorders] = None
    alignment: Optional[CellAlignment] = None
    number_format: Optional[str] = None  # Format code like "0.00" or "yyyy-mm-dd"


class ExcelCellJSON(BaseModel):
    """A single cell in an Excel worksheet.
    
    Captures value, formula, style, and position information.
    """
    id: str  # Unique identifier e.g. "sheet1-A1"
    ref: str  # Cell reference e.g. "A1", "B2"
    row: int  # 1-indexed row number
    col: int  # 1-indexed column number
    col_letter: str  # Column letter e.g. "A", "B", "AA"
    
    # Value
    value: Optional[Any] = None  # The actual cell value (string, number, bool, etc.)
    raw_value: Optional[str] = None  # Raw value from XML (for shared strings, this is the index)
    data_type: Optional[CellDataType] = None
    
    # Formula
    formula: Optional[str] = None  # Formula if present (without '=')
    formula_type: Optional[str] = None  # "normal", "shared", "array"
    shared_formula_ref: Optional[str] = None  # For shared formula master cell
    shared_formula_si: Optional[int] = None  # Shared formula index
    
    # Style
    style: Optional[CellStyle] = None
    style_index: Optional[int] = None  # Index into styles.xml for reconstruction
    
    # Merge info (if this cell is part of a merge)
    is_merged: bool = False
    merge_range: Optional[str] = None  # e.g. "A1:C3" if this cell is part of a merge
    is_merge_origin: bool = False  # True if this is the top-left cell of a merge
    
    # Edit tracking (use regular fields, not private)
    dirty: bool = Field(default=False, exclude=True)  # True if value has been modified
    original_value: Optional[Any] = Field(default=None, exclude=True)  # Original value


class MergedCellRange(BaseModel):
    """A merged cell range in a worksheet."""
    id: str
    ref: str  # Range reference e.g. "B2:F6"
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    start_cell_ref: str  # e.g. "B2"
    end_cell_ref: str  # e.g. "F6"


class DataValidationRule(BaseModel):
    """Data validation rule (dropdowns, input constraints)."""
    id: str
    sqref: str  # Cell reference(s) this applies to e.g. "D20" or "A1:A10"
    validation_type: str  # "list", "whole", "decimal", "date", "time", "textLength", "custom"
    
    # For list (dropdown)
    formula1: Optional[str] = None  # List source: "1,2,3" or named range or formula
    formula2: Optional[str] = None  # For between/notBetween
    
    # Options
    allow_blank: bool = True
    show_input_message: bool = False
    show_error_message: bool = False
    input_title: Optional[str] = None
    input_message: Optional[str] = None
    error_title: Optional[str] = None
    error_message: Optional[str] = None
    error_style: Optional[str] = None  # "stop", "warning", "information"
    
    # Operator for numeric validations
    operator: Optional[str] = None  # "between", "notBetween", "equal", "notEqual", "lessThan", etc.
    
    # Parsed list options (if type is "list")
    options: List[str] = []


class ExcelComment(BaseModel):
    """A cell comment/note."""
    id: str
    cell_ref: str  # Cell this comment is attached to
    author: Optional[str] = None
    text: str
    # Position/size info from VML
    is_visible: bool = False


# =============================================================================
# HYPERLINKS
# =============================================================================

class ExcelHyperlink(BaseModel):
    """A hyperlink in a cell."""
    id: str
    cell_ref: str  # Cell reference e.g. "A1"
    target: Optional[str] = None  # External URL or file path
    location: Optional[str] = None  # Internal location (e.g., "Sheet2!A1")
    display: Optional[str] = None  # Display text
    tooltip: Optional[str] = None  # Tooltip on hover
    r_id: Optional[str] = None  # Relationship ID for external links


# =============================================================================
# CONDITIONAL FORMATTING
# =============================================================================

class ConditionalFormatRule(BaseModel):
    """A single conditional format rule."""
    id: str
    type: str  # "cellIs", "colorScale", "dataBar", "iconSet", "expression", etc.
    priority: int = 1
    operator: Optional[str] = None  # "lessThan", "greaterThan", "equal", "between", etc.
    
    # For formula-based rules
    formula1: Optional[str] = None
    formula2: Optional[str] = None
    
    # Formatting to apply
    dxf_id: Optional[int] = None  # Differential format ID
    
    # Color scale specific
    color_scale: Optional[Dict[str, Any]] = None  # {min_color, mid_color, max_color, min_type, etc.}
    
    # Data bar specific
    data_bar: Optional[Dict[str, Any]] = None  # {min_length, max_length, color, etc.}
    
    # Icon set specific
    icon_set: Optional[Dict[str, Any]] = None  # {icon_set_name, show_value, reverse, etc.}
    
    # Stop if true
    stop_if_true: bool = False


class ConditionalFormatting(BaseModel):
    """Conditional formatting for a range."""
    id: str
    sqref: str  # Cell range(s) e.g. "A1:A10" or "A1:A10 B1:B10"
    rules: List[ConditionalFormatRule] = []


# =============================================================================
# FORM CONTROLS
# =============================================================================

class FormControlType(str, Enum):
    """Types of form controls in Excel."""
    CHECKBOX = "checkbox"
    RADIO = "radio"
    BUTTON = "button"
    DROPDOWN = "dropdown"
    LISTBOX = "listbox"
    SPINNER = "spinner"
    SCROLLBAR = "scrollbar"
    GROUPBOX = "groupbox"
    LABEL = "label"


class FormControl(BaseModel):
    """A form control (checkbox, button, etc.)."""
    id: str
    name: Optional[str] = None
    control_type: FormControlType
    
    # Position (from VML or drawing)
    anchor_type: str = "twoCellAnchor"
    from_col: Optional[int] = None
    from_row: Optional[int] = None
    from_col_off: Optional[int] = None  # Offset in EMUs
    from_row_off: Optional[int] = None
    to_col: Optional[int] = None
    to_row: Optional[int] = None
    to_col_off: Optional[int] = None
    to_row_off: Optional[int] = None
    
    # State
    checked: Optional[bool] = None  # For checkbox/radio
    value: Optional[str] = None  # Current value
    
    # Linked cell (for data binding)
    linked_cell: Optional[str] = None  # Cell reference that stores the value
    
    # For dropdowns/listboxes
    input_range: Optional[str] = None  # Range of options
    selected_index: Optional[int] = None
    
    # For spinners/scrollbars
    min_value: Optional[int] = None
    max_value: Optional[int] = None
    increment: Optional[int] = None
    page_increment: Optional[int] = None
    
    # Visual
    alt_text: Optional[str] = None
    print_object: bool = True
    disabled: bool = False
    
    # VML shape info for preservation
    vml_shape_id: Optional[str] = None


# =============================================================================
# DEFINED NAMES
# =============================================================================

class DefinedName(BaseModel):
    """A defined name (named range, constant, or formula)."""
    id: str
    name: str  # The name (e.g., "PrintArea", "MyRange")
    value: str  # The formula/range (e.g., "Sheet1!$A$1:$D$10")
    
    # Scope
    local_sheet_id: Optional[int] = None  # None = workbook-level, int = sheet-specific
    
    # Properties
    hidden: bool = False
    comment: Optional[str] = None
    
    # Built-in names
    is_builtin: bool = False  # True for _xlnm.Print_Area, _xlnm.Print_Titles, etc.


# =============================================================================
# SHEET VIEW
# =============================================================================

class FreezePane(BaseModel):
    """Freeze pane configuration."""
    x_split: int = 0  # Columns frozen from left
    y_split: int = 0  # Rows frozen from top
    top_left_cell: Optional[str] = None  # First unfrozen cell
    active_pane: str = "bottomRight"  # Which pane is active


class SheetView(BaseModel):
    """Sheet view settings."""
    id: str
    view_type: str = "normal"  # "normal", "pageBreakPreview", "pageLayout"
    zoom_scale: int = 100
    zoom_scale_normal: Optional[int] = None
    zoom_scale_page_layout_view: Optional[int] = None
    
    # Selection
    show_gridlines: bool = True
    show_row_col_headers: bool = True
    show_formulas: bool = False
    show_zeros: bool = True
    
    # Active cell
    active_cell: Optional[str] = None
    active_cell_id: Optional[int] = None
    
    # Freeze pane
    freeze_pane: Optional[FreezePane] = None
    
    # Split (for non-freeze splits)
    split_horizontal: Optional[int] = None  # Split position in twips
    split_vertical: Optional[int] = None


# =============================================================================
# TABLES (ListObjects)
# =============================================================================

class TableColumn(BaseModel):
    """A column in a structured table."""
    id: int
    name: str
    data_cell_style: Optional[str] = None
    header_cell_style: Optional[str] = None
    totals_cell_style: Optional[str] = None
    totals_row_function: Optional[str] = None  # "sum", "count", "average", etc.
    totals_row_formula: Optional[str] = None


class ExcelTable(BaseModel):
    """A structured table (ListObject) in a worksheet."""
    id: str
    name: str  # Table name (unique in workbook)
    display_name: str
    ref: str  # Range reference e.g. "A1:D10"
    
    # Table parts
    header_row_count: int = 1  # Usually 1
    totals_row_count: int = 0  # 0 or 1
    
    # Columns
    columns: List[TableColumn] = []
    
    # Style
    table_style_name: Optional[str] = None  # e.g., "TableStyleMedium2"
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = True
    show_column_stripes: bool = False
    
    # Auto filter
    auto_filter_ref: Optional[str] = None
    
    # Relationship
    r_id: Optional[str] = None


# =============================================================================
# SPARKLINES
# =============================================================================

class Sparkline(BaseModel):
    """A sparkline (mini chart in cell)."""
    id: str
    location: str  # Cell where sparkline appears
    data_range: str  # Source data range
    sparkline_type: str = "line"  # "line", "column", "stacked"
    
    # Colors
    color_series: Optional[str] = None
    color_negative: Optional[str] = None
    color_axis: Optional[str] = None
    color_markers: Optional[str] = None
    color_first: Optional[str] = None
    color_last: Optional[str] = None
    color_high: Optional[str] = None
    color_low: Optional[str] = None
    
    # Options
    show_markers: bool = False
    show_high_point: bool = False
    show_low_point: bool = False
    show_first_point: bool = False
    show_last_point: bool = False
    show_negative_points: bool = False
    
    # Axis
    min_axis_type: str = "individual"  # "individual", "group", "custom"
    max_axis_type: str = "individual"
    right_to_left: bool = False


class SparklineGroup(BaseModel):
    """A group of sparklines sharing settings."""
    id: str
    sparklines: List[Sparkline] = []
    sparkline_type: str = "line"
    
    # Shared settings
    display_empty_cells_as: str = "gap"  # "gap", "zero", "span"
    date_axis: bool = False


class ExcelImage(BaseModel):
    """An embedded image in the worksheet."""
    id: str
    name: Optional[str] = None
    description: Optional[str] = None
    
    # Position (anchor)
    anchor_type: str = "twoCellAnchor"  # or "oneCellAnchor", "absoluteAnchor"
    from_col: Optional[int] = None
    from_row: Optional[int] = None
    to_col: Optional[int] = None
    to_row: Optional[int] = None
    
    # Size
    width_emu: Optional[int] = None  # In EMUs (914400 EMUs = 1 inch)
    height_emu: Optional[int] = None
    
    # Source
    media_path: str  # Path in the XLSX archive e.g. "xl/media/image1.png"
    content_type: Optional[str] = None  # MIME type
    
    # Relationship
    r_id: Optional[str] = None  # Relationship ID


class ColumnInfo(BaseModel):
    """Column dimension/formatting info."""
    min_col: int
    max_col: int
    width: Optional[float] = None
    hidden: bool = False
    best_fit: bool = False
    custom_width: bool = False
    style_index: Optional[int] = None


class RowInfo(BaseModel):
    """Row dimension/formatting info."""
    row: int
    height: Optional[float] = None
    hidden: bool = False
    custom_height: bool = False
    style_index: Optional[int] = None


class ExcelSheetJSON(BaseModel):
    """A single worksheet in an Excel workbook.
    
    Contains cells, merged ranges, validations, images, and comments.
    """
    id: str  # Unique identifier
    name: str  # Sheet tab name
    sheet_index: int  # 0-indexed position in workbook
    
    # Sheet state
    is_hidden: bool = False
    
    # Dimension
    dimension: Optional[str] = None  # Used range e.g. "A1:F20"
    
    # Data
    cells: List[ExcelCellJSON] = []  # All non-empty cells
    
    # Structure
    merged_cells: List[MergedCellRange] = []
    data_validations: List[DataValidationRule] = []
    
    # Columns and rows with custom properties
    columns: List[ColumnInfo] = []
    rows: List[RowInfo] = []
    
    # Embedded content
    images: List[ExcelImage] = []
    comments: List[ExcelComment] = []
    
    # NEW: Complex elements
    hyperlinks: List["ExcelHyperlink"] = []  # Cell hyperlinks
    conditional_formatting: List["ConditionalFormatting"] = []  # CF rules
    form_controls: List["FormControl"] = []  # Checkboxes, buttons, etc.
    tables: List["ExcelTable"] = []  # Structured tables
    sparkline_groups: List["SparklineGroup"] = []  # Sparklines
    
    # Sheet view
    sheet_view: Optional["SheetView"] = None
    
    def get_cell(self, ref: str) -> Optional[ExcelCellJSON]:
        """Get a cell by reference (e.g., 'A1')."""
        for cell in self.cells:
            if cell.ref == ref:
                return cell
        return None
    
    def get_hyperlink(self, ref: str) -> Optional["ExcelHyperlink"]:
        """Get a hyperlink by cell reference."""
        for hl in self.hyperlinks:
            if hl.cell_ref == ref:
                return hl
        return None
    
    def get_table(self, name: str) -> Optional["ExcelTable"]:
        """Get a table by name."""
        for table in self.tables:
            if table.name == name or table.display_name == name:
                return table
        return None


class SharedStringItem(BaseModel):
    """A shared string entry."""
    index: int
    text: str
    rich_text: bool = False  # True if contains formatting


class StyleInfo(BaseModel):
    """Extracted style information for reconstruction."""
    fonts: List[Dict[str, Any]] = []
    fills: List[Dict[str, Any]] = []
    borders: List[Dict[str, Any]] = []
    cell_xfs: List[Dict[str, Any]] = []  # Cell format cross-references
    number_formats: Dict[int, str] = {}  # numFmtId -> formatCode


class ExcelWorkbookJSON(BaseModel):
    """Top-level editable representation of an XLSX workbook.
    
    This is an edit overlay bound to the underlying OOXML, not a full layout model.
    Preserves structure for high-fidelity round-tripping.
    """
    id: str
    filename: Optional[str] = None
    
    # Sheets
    sheets: List[ExcelSheetJSON] = []
    active_sheet_index: int = 0
    
    # Shared data (for reconstruction)
    shared_strings: List[SharedStringItem] = []
    
    # NEW: Workbook-level elements
    defined_names: List["DefinedName"] = []  # Named ranges, print areas, etc.
    
    # Metadata
    created: Optional[str] = None
    modified: Optional[str] = None
    creator: Optional[str] = None
    last_modified_by: Optional[str] = None
    
    def get_sheet(self, name: str) -> Optional[ExcelSheetJSON]:
        """Get a sheet by name."""
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None
    
    def get_sheet_by_index(self, index: int) -> Optional[ExcelSheetJSON]:
        """Get a sheet by index."""
        if 0 <= index < len(self.sheets):
            return self.sheets[index]
        return None
    
    def get_defined_name(self, name: str) -> Optional["DefinedName"]:
        """Get a defined name by name."""
        for dn in self.defined_names:
            if dn.name == name:
                return dn
        return None
