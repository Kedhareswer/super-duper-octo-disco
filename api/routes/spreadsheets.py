"""API routes for Excel/XLSX document processing.

Similar to documents.py but for spreadsheets:
- Upload XLSX -> parse to JSON
- Edit cells
- Export back to XLSX
"""
from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel

from services.excel_engine import (
    xlsx_to_json,
    apply_json_to_xlsx,
    ExcelWorkbookJSON,
)


router = APIRouter(prefix="/spreadsheets", tags=["spreadsheets"])

# In-memory storage for active spreadsheets (similar to documents.py)
_active_spreadsheets: dict[str, ExcelWorkbookJSON] = {}
_spreadsheet_paths: dict[str, str] = {}

UPLOAD_DIR = Path(__file__).parent.parent.parent / "data" / "uploads"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "data" / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# MODELS
# =============================================================================

class CellEditRequest(BaseModel):
    """Request to edit a cell value."""
    sheet: str  # Sheet name or index
    cell: str  # Cell reference e.g. "A1"
    value: str | int | float | bool | None


class BatchCellEditRequest(BaseModel):
    """Request to edit multiple cells."""
    edits: list[CellEditRequest]


class AIEditCellRequest(BaseModel):
    """Request for AI-driven edit on a cell."""
    sheet: str  # Sheet name or index
    cell: str  # Cell reference e.g. "A1"
    instruction: str  # AI instruction


class CheckboxUpdateRequest(BaseModel):
    """Request to update a checkbox state in Excel."""
    sheet: str  # Sheet name or index
    control_id: str  # Form control ID
    checked: bool


class DropdownUpdateRequest(BaseModel):
    """Request to update a dropdown/data validation selection in Excel."""
    sheet: str  # Sheet name or index
    cell: str  # Cell reference e.g. "A1"
    value: str  # Selected value


# =============================================================================
# ENDPOINTS
# =============================================================================

@router.post("/", response_model=dict)
async def upload_spreadsheet(file: UploadFile = File(...)):
    """Upload an XLSX file and parse it to JSON structure.
    
    Returns a simplified summary suitable for the UI, not the full workbook.
    """
    if not file.filename:
        raise HTTPException(400, "No filename provided")
    
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Only .xlsx files are supported")
    
    # Save uploaded file
    spreadsheet_id = f"{uuid.uuid4().hex[:8]}_{file.filename}"
    file_path = UPLOAD_DIR / spreadsheet_id
    
    content = await file.read()
    file_path.write_bytes(content)
    
    try:
        # Parse to JSON
        workbook = xlsx_to_json(str(file_path), spreadsheet_id)
        
        # Store in memory
        _active_spreadsheets[spreadsheet_id] = workbook
        _spreadsheet_paths[spreadsheet_id] = str(file_path)
        
        # Return summary for UI
        return _workbook_to_ui_summary(workbook, spreadsheet_id)
    
    except Exception as e:
        # Clean up on error
        if file_path.exists():
            file_path.unlink()
        raise HTTPException(500, f"Failed to parse spreadsheet: {e}")


@router.get("/{spreadsheet_id}")
async def get_spreadsheet(spreadsheet_id: str):
    """Get the current state of a spreadsheet."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


@router.put("/{spreadsheet_id}")
async def update_spreadsheet(spreadsheet_id: str, data: dict):
    """Update spreadsheet from UI (full JSON sync)."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    # For now, just update the in-memory workbook with cell changes
    workbook = _active_spreadsheets[spreadsheet_id]
    
    # Apply any cell edits from the data
    if "sheets" in data:
        for sheet_data in data["sheets"]:
            sheet = workbook.get_sheet(sheet_data.get("name", ""))
            if not sheet:
                continue
            
            for cell_data in sheet_data.get("cells", []):
                cell = sheet.get_cell(cell_data.get("ref", ""))
                if cell:
                    cell.value = cell_data.get("value")
    
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


@router.post("/{spreadsheet_id}/cell")
async def edit_cell(spreadsheet_id: str, edit: CellEditRequest):
    """Edit a single cell value.
    
    If the cell contains a formula, the formula will be cleared and replaced
    with the new value. A warning is returned in the response.
    """
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    
    # Find sheet
    sheet = workbook.get_sheet(edit.sheet)
    if not sheet:
        # Try by index
        try:
            sheet = workbook.get_sheet_by_index(int(edit.sheet))
        except (ValueError, TypeError):
            pass
    
    if not sheet:
        raise HTTPException(404, f"Sheet not found: {edit.sheet}")
    
    # Track if we're overwriting a formula
    formula_warning = None
    
    # Find or create cell
    cell = sheet.get_cell(edit.cell)
    if cell:
        # Check if this cell has a formula
        if cell.formula:
            formula_warning = f"Cell {edit.cell} had formula '={cell.formula}' which was cleared"
            # Clear the formula since we're setting a static value
            cell.formula = None
            cell.formula_type = None
            cell.shared_formula_ref = None
            cell.shared_formula_si = None
        
        # Track original value for dirty detection
        if cell.original_value is None:
            cell.original_value = cell.value
        cell.value = edit.value
        cell.dirty = True
    else:
        # Cell doesn't exist - create it
        from services.excel_engine.parser import parse_cell_ref
        col_letter, col, row = parse_cell_ref(edit.cell)
        from services.excel_engine.schemas import ExcelCellJSON
        new_cell = ExcelCellJSON(
            id=f"{sheet.id}-{edit.cell}",
            ref=edit.cell,
            row=row,
            col=col,
            col_letter=col_letter,
            value=edit.value,
            dirty=True,
        )
        sheet.cells.append(new_cell)
    
    result = _workbook_to_ui_summary(workbook, spreadsheet_id)
    
    # Add warning if formula was cleared
    if formula_warning:
        result["warnings"] = [formula_warning]
    
    return result


@router.post("/{spreadsheet_id}/cells")
async def edit_cells(spreadsheet_id: str, batch: BatchCellEditRequest):
    """Edit multiple cells at once."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    
    for edit in batch.edits:
        sheet = workbook.get_sheet(edit.sheet)
        if not sheet:
            try:
                sheet = workbook.get_sheet_by_index(int(edit.sheet))
            except (ValueError, TypeError):
                continue
        
        if sheet:
            cell = sheet.get_cell(edit.cell)
            if cell:
                cell.value = edit.value
    
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


@router.post("/{spreadsheet_id}/export/file")
async def export_spreadsheet(spreadsheet_id: str):
    """Export spreadsheet back to XLSX file."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    base_path = _spreadsheet_paths.get(spreadsheet_id)
    
    if not base_path or not Path(base_path).exists():
        raise HTTPException(500, "Original spreadsheet file not found")
    
    # Generate output path
    output_filename = spreadsheet_id.replace(".xlsx", "_copy.xlsx").replace(".XLSX", "_copy.xlsx")
    if not output_filename.endswith(".xlsx"):
        output_filename += "_copy.xlsx"
    output_path = OUTPUT_DIR / output_filename
    
    try:
        apply_json_to_xlsx(workbook, base_path, str(output_path))
        
        return FileResponse(
            path=str(output_path),
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        raise HTTPException(500, f"Export failed: {e}")


@router.post("/{spreadsheet_id}/ai-edit")
async def ai_edit_cell(spreadsheet_id: str, payload: AIEditCellRequest):
    """Apply an AI edit to a cell value.
    
    Uses LangGraph + Gemini as the AI backend.
    """
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    
    # Find sheet
    sheet = workbook.get_sheet(payload.sheet)
    if not sheet:
        try:
            sheet = workbook.get_sheet_by_index(int(payload.sheet))
        except (ValueError, TypeError):
            pass
    
    if not sheet:
        raise HTTPException(404, f"Sheet not found: {payload.sheet}")
    
    # Find cell - create if doesn't exist
    cell = sheet.get_cell(payload.cell)
    
    # Get current cell value as text
    original_text = ""
    if cell and cell.value is not None:
        original_text = str(cell.value)
    
    # If cell is empty, use the instruction as a generation prompt instead
    if not original_text.strip():
        # For empty cells, we'll generate content based on the instruction
        original_text = "[EMPTY CELL - Generate content]"
    
    # Import AI agent
    from services.ai_agent import DocumentEditAgent
    
    try:
        agent = DocumentEditAgent()
        result = await agent.edit(
            text=original_text,
            instruction=payload.instruction,
            context=f"Excel cell {payload.cell} in sheet '{sheet.name}'"
        )
        edited_text = result.get("edited_text", original_text)
        
        # Apply edit to cell
        if cell:
            if cell.original_value is None:
                cell.original_value = cell.value
            cell.value = edited_text
            cell.dirty = True
        else:
            # Create new cell if it didn't exist
            from services.excel_engine.parser import parse_cell_ref
            from services.excel_engine.schemas import ExcelCellJSON
            col_letter, col, row = parse_cell_ref(payload.cell)
            new_cell = ExcelCellJSON(
                id=f"{sheet.id}-{payload.cell}",
                ref=payload.cell,
                row=row,
                col=col,
                col_letter=col_letter,
                value=edited_text,
                dirty=True,
            )
            sheet.cells.append(new_cell)
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"AI edit failed: {e}")
    
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


@router.post("/{spreadsheet_id}/checkbox")
async def update_checkbox(spreadsheet_id: str, payload: CheckboxUpdateRequest):
    """Update a checkbox (form control) state in the spreadsheet."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    
    # Find sheet
    sheet = workbook.get_sheet(payload.sheet)
    if not sheet:
        try:
            sheet = workbook.get_sheet_by_index(int(payload.sheet))
        except (ValueError, TypeError):
            pass
    
    if not sheet:
        raise HTTPException(404, f"Sheet not found: {payload.sheet}")
    
    # Find form control
    control_found = False
    for control in sheet.form_controls:
        if control.id == payload.control_id:
            control.checked = payload.checked
            control_found = True
            
            # Also update linked cell if present
            if control.linked_cell:
                linked_cell = sheet.get_cell(control.linked_cell)
                if linked_cell:
                    if linked_cell.original_value is None:
                        linked_cell.original_value = linked_cell.value
                    # Excel uses TRUE/FALSE for checkbox linked cells
                    linked_cell.value = payload.checked
                    linked_cell.dirty = True
            break
    
    if not control_found:
        raise HTTPException(404, f"Checkbox control not found: {payload.control_id}")
    
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


@router.post("/{spreadsheet_id}/dropdown")
async def update_dropdown(spreadsheet_id: str, payload: DropdownUpdateRequest):
    """Update a cell value that has data validation (dropdown)."""
    if spreadsheet_id not in _active_spreadsheets:
        raise HTTPException(404, "Spreadsheet not found")
    
    workbook = _active_spreadsheets[spreadsheet_id]
    
    # Find sheet
    sheet = workbook.get_sheet(payload.sheet)
    if not sheet:
        try:
            sheet = workbook.get_sheet_by_index(int(payload.sheet))
        except (ValueError, TypeError):
            pass
    
    if not sheet:
        raise HTTPException(404, f"Sheet not found: {payload.sheet}")
    
    # Find the data validation for this cell
    validation = None
    cell_ref_upper = payload.cell.upper()
    for v in sheet.data_validations:
        if not v.sqref:
            continue
        # Check if cell is in the sqref range
        ranges = str(v.sqref).split()
        for r in ranges:
            if ":" in r:
                # Range like A1:A10
                start, end = r.split(":", 1)
                # Simple check - could be more robust
                if _cell_in_range(cell_ref_upper, start.upper(), end.upper()):
                    validation = v
                    break
            else:
                # Single cell
                if r.upper() == cell_ref_upper:
                    validation = v
                    break
        if validation:
            break
    
    # Validate the value if we found a list validation
    if validation and validation.validation_type == "list" and validation.options:
        if payload.value not in validation.options:
            raise HTTPException(400, f"Invalid selection. Must be one of: {validation.options}")
    
    # Find or create cell
    cell = sheet.get_cell(payload.cell)
    if cell:
        if cell.original_value is None:
            cell.original_value = cell.value
        cell.value = payload.value
        cell.dirty = True
    else:
        # Create new cell
        from services.excel_engine.parser import parse_cell_ref
        from services.excel_engine.schemas import ExcelCellJSON
        col_letter, col, row = parse_cell_ref(payload.cell)
        new_cell = ExcelCellJSON(
            id=f"{sheet.id}-{payload.cell}",
            ref=payload.cell,
            row=row,
            col=col,
            col_letter=col_letter,
            value=payload.value,
            dirty=True,
        )
        sheet.cells.append(new_cell)
    
    return _workbook_to_ui_summary(workbook, spreadsheet_id)


def _cell_in_range(cell_ref: str, start_ref: str, end_ref: str) -> bool:
    """Check if a cell reference is within a range."""
    from services.excel_engine.parser import parse_cell_ref
    try:
        _, cell_col, cell_row = parse_cell_ref(cell_ref)
        _, start_col, start_row = parse_cell_ref(start_ref)
        _, end_col, end_row = parse_cell_ref(end_ref)
        
        # Normalize range bounds
        min_row, max_row = min(start_row, end_row), max(start_row, end_row)
        min_col, max_col = min(start_col, end_col), max(start_col, end_col)
        
        return min_row <= cell_row <= max_row and min_col <= cell_col <= max_col
    except Exception:
        return False


# =============================================================================
# HELPERS
# =============================================================================

def _workbook_to_ui_summary(workbook: ExcelWorkbookJSON, spreadsheet_id: str) -> dict:
    """Convert workbook to a UI-friendly summary structure."""
    sheets = []
    
    for sheet in workbook.sheets:
        # Build quick lookup for data validations (dropdowns) by cell ref
        validation_by_cell: dict[str, dict] = {}
        for v in sheet.data_validations:
            if not v.sqref:
                continue
            # sqref can contain multiple ranges/refs separated by spaces
            ranges = str(v.sqref).split()
            for r in ranges:
                if ":" in r:
                    start, end = r.split(":", 1)
                    # Parse refs like A1, B10
                    def _parse_ref(ref: str) -> tuple[int, int] | None:
                        ref = ref.upper()
                        i = 0
                        while i < len(ref) and ref[i].isalpha():
                            i += 1
                        if i == 0 or i == len(ref):
                            return None
                        col_letters = ref[:i]
                        row_str = ref[i:]
                        if not row_str.isdigit():
                            return None
                        row = int(row_str)
                        col = 0
                        for ch in col_letters:
                            col = col * 26 + (ord(ch) - 64)
                        return row, col
                    start_rc = _parse_ref(start)
                    end_rc = _parse_ref(end)
                    if not start_rc or not end_rc:
                        continue
                    (start_row, start_col), (end_row, end_col) = start_rc, end_rc
                    r1, r2 = sorted([start_row, end_row])
                    c1, c2 = sorted([start_col, end_col])
                    for row in range(r1, r2 + 1):
                        for col in range(c1, c2 + 1):
                            # Reconstruct ref like A1
                            tmp_col = col
                            letters = ""
                            while tmp_col > 0:
                                tmp_col, rem = divmod(tmp_col - 1, 26)
                                letters = chr(65 + rem) + letters
                            cell_ref = f"{letters}{row}"
                            validation_by_cell[cell_ref] = {
                                "id": v.id,
                                "sqref": v.sqref,
                                "type": v.validation_type,
                                "options": v.options,
                            }
                else:
                    # Single cell ref
                    cell_ref = r.upper()
                    validation_by_cell[cell_ref] = {
                        "id": v.id,
                        "sqref": v.sqref,
                        "type": v.validation_type,
                        "options": v.options,
                    }

        # Build quick lookup for form controls (e.g., checkboxes) by linked cell ref
        control_by_cell: dict[str, dict] = {}
        for c in sheet.form_controls:
            if not c.linked_cell:
                continue
            cell_ref = str(c.linked_cell).upper()
            control_by_cell[cell_ref] = {
                "id": c.id,
                "type": c.control_type.value,
                "checked": c.checked,
                "linked_cell": c.linked_cell,
            }

        # Get visible cells (limit for UI performance)
        cells = []
        for cell in sheet.cells[:5000]:  # Limit to 5000 cells for UI
            # Build comprehensive style info
            style_info = None
            if cell.style:
                style_info = {
                    "bold": cell.style.font.bold if cell.style.font else False,
                    "italic": cell.style.font.italic if cell.style.font else False,
                    "underline": cell.style.font.underline if cell.style.font else False,
                    "color": cell.style.font.color if cell.style.font else None,
                    "font_size": cell.style.font.size if cell.style.font else None,
                    "bg_color": cell.style.fill.fg_color if cell.style.fill else None,
                    "pattern": cell.style.fill.pattern_type if cell.style.fill else None,
                    "h_align": cell.style.alignment.horizontal if cell.style.alignment else None,
                    "v_align": cell.style.alignment.vertical if cell.style.alignment else None,
                    "wrap": cell.style.alignment.wrap_text if cell.style.alignment else False,
                    "borders": {
                        "left": cell.style.borders.left.style if cell.style.borders and cell.style.borders.left else None,
                        "right": cell.style.borders.right.style if cell.style.borders and cell.style.borders.right else None,
                        "top": cell.style.borders.top.style if cell.style.borders and cell.style.borders.top else None,
                        "bottom": cell.style.borders.bottom.style if cell.style.borders and cell.style.borders.bottom else None,
                    } if cell.style.borders else None,
                }
            
            # Attach dropdown metadata if this cell participates in a list validation
            cell_ref_upper = cell.ref.upper() if cell.ref else ""
            dv = validation_by_cell.get(cell_ref_upper)
            dropdown_info = None
            if dv and dv.get("type") == "list":
                dropdown_info = {
                    "validation_id": dv.get("id"),
                    "options": dv.get("options") or [],
                }

            # Attach checkbox metadata if this cell is the linked_cell of a form control
            control = control_by_cell.get(cell_ref_upper)
            checkbox_info = None
            if control and control.get("type") == "checkbox":
                checkbox_info = {
                    "control_id": control.get("id"),
                    "checked": bool(control.get("checked")),
                }

            cells.append({
                "id": cell.id,
                "ref": cell.ref,
                "row": cell.row,
                "col": cell.col,
                "value": cell.value,
                "formula": cell.formula,
                "has_formula": bool(cell.formula),  # Convenience flag for UI
                "formula_type": cell.formula_type,  # "normal", "shared", "array"
                "is_merged": cell.is_merged,
                "merge_range": cell.merge_range,
                "is_merge_origin": cell.is_merge_origin,
                "style": style_info,
                "dropdown": dropdown_info,
                "checkbox": checkbox_info,
            })
        
        # Merged cells info
        merges = [
            {
                "ref": m.ref,
                "start_row": m.start_row,
                "start_col": m.start_col,
                "end_row": m.end_row,
                "end_col": m.end_col,
            }
            for m in sheet.merged_cells
        ]
        
        # Data validations (dropdowns)
        validations = [
            {
                "id": v.id,
                "sqref": v.sqref,
                "type": v.validation_type,
                "options": v.options,
            }
            for v in sheet.data_validations
        ]
        
        # Conditional formatting - expose full details for editing
        conditional_formatting = [
            {
                "id": cf.id,
                "sqref": cf.sqref,  # Cell range(s) affected
                "rules": [
                    {
                        "id": rule.id,
                        "type": rule.type,  # "cellIs", "colorScale", "dataBar", etc.
                        "priority": rule.priority,
                        "operator": rule.operator,  # "lessThan", "greaterThan", etc.
                        "formula1": rule.formula1,
                        "formula2": rule.formula2,
                        "stop_if_true": rule.stop_if_true,
                    }
                    for rule in cf.rules
                ]
            }
            for cf in sheet.conditional_formatting
        ]
        cf_count = len(conditional_formatting)
        
        # Form controls
        controls = [
            {
                "id": c.id,
                "type": c.control_type.value,
                "checked": c.checked,
                "linked_cell": c.linked_cell,
            }
            for c in sheet.form_controls
        ]
        
        # Freeze pane
        freeze = None
        if sheet.sheet_view and sheet.sheet_view.freeze_pane:
            fp = sheet.sheet_view.freeze_pane
            freeze = {
                "rows": fp.y_split,
                "cols": fp.x_split,
            }
        
        sheets.append({
            "id": sheet.id,
            "name": sheet.name,
            "index": sheet.sheet_index,
            "is_hidden": sheet.is_hidden,
            "dimension": sheet.dimension,
            "cells": cells,
            "cell_count": len(sheet.cells),
            "merged_cells": merges,
            "data_validations": validations,
            "conditional_formatting": conditional_formatting,  # Full CF details
            "conditional_formatting_count": cf_count,
            "form_controls": controls,
            "images_count": len(sheet.images),
            "comments_count": len(sheet.comments),
            "hyperlinks_count": len(sheet.hyperlinks),
            "freeze_pane": freeze,
            "zoom": sheet.sheet_view.zoom_scale if sheet.sheet_view else 100,
        })
    
    # Defined names
    defined_names = [
        {
            "name": dn.name,
            "value": dn.value,
            "is_builtin": dn.is_builtin,
        }
        for dn in workbook.defined_names
    ]
    
    return {
        "id": spreadsheet_id,
        "filename": workbook.filename,
        "sheets": sheets,
        "active_sheet_index": workbook.active_sheet_index,
        "defined_names": defined_names,
        "metadata": {
            "created": workbook.created,
            "modified": workbook.modified,
            "creator": workbook.creator,
        },
    }
