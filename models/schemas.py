from __future__ import annotations

from enum import Enum
from typing import List, Literal, Optional, Union

from pydantic import BaseModel


class BlockType(str, Enum):
    PARAGRAPH = "paragraph"
    TABLE = "table"
    DRAWING = "drawing"


class DrawingBlock(BaseModel):
    """A drawing/image placeholder in the document."""
    
    type: BlockType = BlockType.DRAWING
    id: str
    xml_ref: str
    name: Optional[str] = None
    width_inches: float = 0
    height_inches: float = 0
    drawing_type: str = "unknown"  # "vector_group", "image", "chart", etc.


# =============================================================================
# INLINE CONTENT: Runs, Checkboxes, Dropdowns (all can appear inside paragraphs)
# =============================================================================

class TextRun(BaseModel):
    """Inline text run within a paragraph, bound to a specific OOXML <w:r>."""
    
    run_type: Literal["text"] = "text"
    id: str
    xml_ref: str  # positional reference into document.xml (e.g. "p[15]/r[3]")
    text: str
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None  # Hex color e.g. "FF0000" for red


class CheckboxRun(BaseModel):
    """Inline checkbox content control (SDT with w14:checkbox)."""
    
    run_type: Literal["checkbox"] = "checkbox"
    id: str
    xml_ref: str  # reference to the SDT element
    label: Optional[str] = None
    checked: bool = False


class DropdownRun(BaseModel):
    """Inline dropdown/combo content control (SDT with w:dropDownList or w:comboBox)."""
    
    run_type: Literal["dropdown"] = "dropdown"
    id: str
    xml_ref: str  # reference to the SDT element
    label: Optional[str] = None
    options: List[str] = []
    selected: Optional[str] = None


# Union of all inline content types
InlineContent = Union[TextRun, CheckboxRun, DropdownRun]


# Legacy alias for backward compatibility
Run = TextRun


class ParagraphBlock(BaseModel):
    """A paragraph in the document body or inside a table cell.
    
    The `runs` field contains inline content which can be:
    - TextRun: Regular text with formatting
    - CheckboxRun: A checkbox content control
    - DropdownRun: A dropdown/combo content control
    """

    type: BlockType = BlockType.PARAGRAPH
    id: str
    xml_ref: str  # reference to <w:p>
    style_name: Optional[str] = None
    runs: List[InlineContent]  # Can contain TextRun, CheckboxRun, DropdownRun


class CellBorder(BaseModel):
    """Border style for a cell edge."""
    style: str = "none"  # none, single, double, dashed, dotted, etc.
    width: int = 0  # in eighths of a point
    color: Optional[str] = None  # Hex color


class CellBorders(BaseModel):
    """All borders for a cell."""
    top: Optional[CellBorder] = None
    bottom: Optional[CellBorder] = None
    left: Optional[CellBorder] = None
    right: Optional[CellBorder] = None


class TableCell(BaseModel):
    """A logical table cell, potentially spanning multiple rows/columns."""

    id: str
    xml_ref: str  # reference to <w:tc>
    row_span: int = 1
    col_span: int = 1
    background_color: Optional[str] = None  # Hex color from w:shd fill
    borders: Optional[CellBorders] = None  # Cell border styles
    v_merge: Optional[str] = None  # "restart" = start of merge, "continue" = continuation, None = no merge
    # Paragraphs and nested tables contained in the cell
    # Uses forward reference for nested TableBlock to handle circular dependency
    blocks: List[Union[ParagraphBlock, "TableBlock"]] = []


class TableRow(BaseModel):
    """A table row consisting of one or more cells."""

    id: str
    xml_ref: str  # reference to <w:tr>
    cells: List[TableCell]


class TableBlock(BaseModel):
    """A table in the document body."""

    type: BlockType = BlockType.TABLE
    id: str
    xml_ref: str  # reference to <w:tbl>
    rows: List[TableRow]


Block = Union[ParagraphBlock, TableBlock, DrawingBlock]

# Rebuild models to resolve forward references for nested tables
TableCell.model_rebuild()


class CheckboxField(BaseModel):
    """Logical view of a checkbox form field.
    
    DEPRECATED: Use CheckboxRun inline in paragraph runs instead.
    This is kept for backward compatibility only.
    """

    id: str
    xml_ref: str
    label: Optional[str] = None
    checked: bool = False


class DropdownField(BaseModel):
    """Logical view of a dropdown form field.
    
    DEPRECATED: Use DropdownRun inline in paragraph runs instead.
    This is kept for backward compatibility only.
    """

    id: str
    xml_ref: str
    label: Optional[str] = None
    options: List[str] = []
    selected: Optional[str] = None


class DocumentJSON(BaseModel):
    """Top-level editable representation of a DOCX document.

    This is an edit overlay bound to the underlying OOXML, not a full layout model.
    
    Content controls (checkboxes, dropdowns) are now inline in paragraph runs.
    The top-level checkboxes/dropdowns arrays are deprecated but maintained for compatibility.
    """

    id: str
    title: Optional[str] = None
    blocks: List[Block]
    # DEPRECATED: These are now inline in paragraph runs as CheckboxRun/DropdownRun
    # Kept for backward compatibility - populated from inline runs
    checkboxes: List[CheckboxField] = []
    dropdowns: List[DropdownField] = []
    
    def get_all_checkboxes(self) -> List[CheckboxRun]:
        """Extract all inline checkbox controls from the document."""
        result: List[CheckboxRun] = []
        
        def extract_from_blocks(blocks: List[Block]) -> None:
            for block in blocks:
                if isinstance(block, ParagraphBlock):
                    for run in block.runs:
                        if isinstance(run, CheckboxRun):
                            result.append(run)
                elif isinstance(block, TableBlock):
                    for row in block.rows:
                        for cell in row.cells:
                            extract_from_blocks(cell.blocks)
        
        extract_from_blocks(self.blocks)
        return result
    
    def get_all_dropdowns(self) -> List[DropdownRun]:
        """Extract all inline dropdown controls from the document."""
        result: List[DropdownRun] = []
        
        def extract_from_blocks(blocks: List[Block]) -> None:
            for block in blocks:
                if isinstance(block, ParagraphBlock):
                    for run in block.runs:
                        if isinstance(run, DropdownRun):
                            result.append(run)
                elif isinstance(block, TableBlock):
                    for row in block.rows:
                        for cell in row.cells:
                            extract_from_blocks(cell.blocks)
        
        extract_from_blocks(self.blocks)
        return result


class ValidationErrorDetail(BaseModel):
    field: str
    message: str


class ValidationResult(BaseModel):
    is_valid: bool
    errors: List[ValidationErrorDetail] = []
