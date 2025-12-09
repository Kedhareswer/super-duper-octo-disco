from __future__ import annotations

import logging
from pathlib import Path

from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel

from models.schemas import BlockType, DocumentJSON
from services.document_engine import (
    apply_json_to_docx,
    docx_to_json,
    validate_document_json,
)
from services.db import Document, get_session
from services.validation import (
    validate_parse_stage,
    validate_export_stage,
    validate_full_roundtrip,
    extract_json_content,
    extract_raw_docx_content,
)
from services.debug_output import (
    save_docx_structure,
    save_parsed_json,
    save_validation_report,
    save_export_comparison,
    create_debug_manifest,
    get_debug_dir,
)

logger = logging.getLogger(__name__)


router = APIRouter(prefix="/documents", tags=["documents"])


DATA_ROOT = Path("data")
UPLOAD_ROOT = DATA_ROOT / "uploads"
EXPORT_ROOT = DATA_ROOT / "exports"


@router.post("/", response_model=DocumentJSON)
async def upload_document(file: UploadFile = File(...)) -> DocumentJSON:
    """Upload a DOCX and convert it into editable JSON.

    v1: stores file on local filesystem and metadata in the SQLite `documents`
    table via services.db.
    """

    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported in v1")

    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)

    document_id = file.filename  # v1: naive id; replace with UUID in real system
    docx_path = UPLOAD_ROOT / document_id

    content = await file.read()
    docx_path.write_bytes(content)

    # DEBUG: Save DOCX structure (XML files inside the ZIP)
    logger.info(f"[DEBUG] Extracting DOCX structure for {document_id}")
    docx_structure = save_docx_structure(document_id, str(docx_path))

    # Stage 1: Parse DOCX to JSON
    logger.info(f"[PARSE] Starting DOCX→JSON conversion for {document_id}")
    json_doc = docx_to_json(str(docx_path), document_id=document_id)
    
    # DEBUG: Save parsed JSON with statistics
    parse_stats = save_parsed_json(document_id, json_doc)
    
    # Stage 2: Validate parsing
    parse_report = validate_parse_stage(str(docx_path), json_doc)
    
    # DEBUG: Save validation report
    save_validation_report(document_id, "parse", {
        "has_errors": parse_report.has_errors,
        "has_warnings": parse_report.has_warnings,
        "stages": [{"stage": s.stage, "total_chars": s.total_chars, "paragraph_count": s.paragraph_count, 
                    "table_count": s.table_count, "checkbox_count": s.checkbox_count} for s in parse_report.stages],
        "issues": [{"severity": i.severity, "category": i.category, "message": i.message} for i in parse_report.issues],
    })
    
    # DEBUG: Create manifest of all debug files
    create_debug_manifest(document_id)
    
    if parse_report.has_errors:
        logger.error(f"[PARSE] Validation errors for {document_id}: {[i.message for i in parse_report.issues if i.severity == 'error']}")
    else:
        logger.info(f"[PARSE] Validation passed for {document_id}: {parse_report.stages[-1].total_chars} chars, {parse_report.stages[-1].paragraph_count} paragraphs")
    
    logger.info(f"[DEBUG] Debug files saved to: data/debug/{document_id}/")

    # Persist in DB (upsert semantics for simplicity)
    with get_session() as db:
        existing = db.get(Document, document_id)
        if existing is None:
            doc_row = Document(
                id=document_id,
                base_docx_path=str(docx_path),
                json=json_doc.model_dump_json(),
                version=1,
                latest_export_path=None,
            )
            db.add(doc_row)
        else:
            existing.base_docx_path = str(docx_path)
            existing.json = json_doc.model_dump_json()
            existing.version = 1
            existing.latest_export_path = None

    return json_doc


@router.get("/{document_id}", response_model=DocumentJSON)
async def get_document(document_id: str) -> DocumentJSON:
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")
        return DocumentJSON.model_validate_json(row.json)


@router.put("/{document_id}", response_model=DocumentJSON)
async def update_document(document_id: str, updated: DocumentJSON) -> DocumentJSON:
    if updated.id != document_id:
        raise HTTPException(status_code=400, detail="Document id in body does not match path")

    validation = validate_document_json(updated)
    if not validation.is_valid:
        raise HTTPException(
            status_code=400,
            detail={
                "message": "Validation failed",
                "errors": [e.dict() for e in validation.errors],
            },
        )

    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")
        row.json = updated.model_dump_json()
        row.version = (row.version or 1) + 1

    return updated


class ExportResponse(DocumentJSON):
    """Response when exporting a document, including path to generated DOCX."""

    export_path: str
    version: int


@router.post("/{document_id}/export", response_model=ExportResponse)
async def export_document(document_id: str) -> ExportResponse:
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")

        json_doc = DocumentJSON.model_validate_json(row.json)
        
        # Get original JSON for comparison
        original_json = docx_to_json(row.base_docx_path, document_id)

        EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
        next_version = (row.version or 1) + 1
        out_path = EXPORT_ROOT / f"{document_id}.v{next_version}.docx"

        # Stage 1: Export JSON to DOCX
        logger.info(f"[EXPORT] Starting JSON→DOCX export for {document_id} v{next_version}")
        result_path = apply_json_to_docx(
            json_doc=json_doc,
            base_docx_path=row.base_docx_path,
            out_docx_path=str(out_path),
        )

        # Stage 2: Validate export
        export_report = validate_export_stage(original_json, json_doc, row.base_docx_path, result_path)
        
        # DEBUG: Save export comparison
        save_export_comparison(document_id, row.base_docx_path, result_path)
        save_validation_report(document_id, "export", {
            "has_errors": export_report.has_errors,
            "has_warnings": export_report.has_warnings,
            "stages": [{"stage": s.stage, "total_chars": s.total_chars, "paragraph_count": s.paragraph_count} for s in export_report.stages],
            "issues": [{"severity": i.severity, "category": i.category, "message": i.message} for i in export_report.issues],
        })
        create_debug_manifest(document_id)
        
        if export_report.has_errors:
            logger.error(f"[EXPORT] Validation errors for {document_id}: {[i.message for i in export_report.issues if i.severity == 'error']}")
        else:
            logger.info(f"[EXPORT] Validation passed for {document_id}: exported to {result_path}")

        row.latest_export_path = result_path
        row.version = next_version

    # Compose response that includes the document JSON plus export metadata
    response_data = json_doc.model_dump()
    return ExportResponse(**response_data, export_path=result_path, version=next_version)


@router.post("/{document_id}/export/file")
async def download_export_document(document_id: str):
    """Export and return the DOCX file itself.

    This is primarily for backend/tests and curl usage. The existing
    `/documents/{id}/export` endpoint continues to return JSON metadata for
    the frontend.
    """

    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")

        json_doc = DocumentJSON.model_validate_json(row.json)

        EXPORT_ROOT.mkdir(parents=True, exist_ok=True)
        next_version = (row.version or 1) + 1
        out_path = EXPORT_ROOT / f"{document_id}.v{next_version}.docx"

        result_path = apply_json_to_docx(
            json_doc=json_doc,
            base_docx_path=row.base_docx_path,
            out_docx_path=str(out_path),
        )

        row.latest_export_path = result_path
        row.version = next_version

    return FileResponse(
        path=result_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=Path(result_path).name,
    )


def _render_html_from_document(doc: DocumentJSON) -> str:
    """Render a minimal but readable HTML representation of a DocumentJSON.

    This is intentionally simple and semantic (not pixel-perfect):
    - Paragraphs become <p> with inline <strong>/<em> for bold/italic.
    - Tables become <table> with <tr>/<td> populated by cell paragraph texts.
    """

    parts: list[str] = [
        "<html><head><meta charset='utf-8'><title>Document Preview</title>",
        "<style>body{font-family:system-ui, sans-serif;padding:16px;}table{border-collapse:collapse;margin:12px 0;width:100%;}td,th{border:1px solid #ddd;padding:4px 6px;vertical-align:top;}</style>",
        "</head><body>",
    ]

    for block in doc.blocks:
        if getattr(block, "type", None) == BlockType.PARAGRAPH:
            # Combine runs into HTML spans
            run_html_parts: list[str] = []
            for run in block.runs:
                text = (run.text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                if not text:
                    continue
                if run.bold:
                    text = f"<strong>{text}</strong>"
                if run.italic:
                    text = f"<em>{text}</em>"
                run_html_parts.append(text)
            paragraph_html = "".join(run_html_parts) or "&nbsp;"
            parts.append(f"<p>{paragraph_html}</p>")

        elif getattr(block, "type", None) == BlockType.TABLE:
            def render_table_html(table_block) -> str:
                """Recursively render a table and its nested tables to HTML."""
                html_parts = ["<table>"]
                for row in table_block.rows:
                    html_parts.append("<tr>")
                    for cell in row.cells:
                        # Concatenate all paragraph texts and nested tables inside the cell
                        cell_content_parts: list[str] = []
                        for cell_block in cell.blocks:
                            if getattr(cell_block, "type", None) == BlockType.PARAGRAPH:
                                for run in cell_block.runs:
                                    t = (run.text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                    cell_content_parts.append(t)
                            elif getattr(cell_block, "type", None) == BlockType.TABLE:
                                # Nested table - render recursively
                                cell_content_parts.append(render_table_html(cell_block))
                        cell_content = " ".join(p for p in cell_content_parts if p) or "&nbsp;"
                        html_parts.append(f"<td>{cell_content}</td>")
                    html_parts.append("</tr>")
                html_parts.append("</table>")
                return "".join(html_parts)
            
            parts.append(render_table_html(block))

    parts.append("</body></html>")
    return "".join(parts)


@router.get("/{document_id}/preview/html", response_class=HTMLResponse)
async def preview_document_html(document_id: str) -> HTMLResponse:
    """Return a simple HTML preview of the document for in-app viewing."""

    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")
        json_doc = DocumentJSON.model_validate_json(row.json)

    html = _render_html_from_document(json_doc)
    return HTMLResponse(content=html)


class AIEditRequest(BaseModel):
    """Request for an AI-driven edit on a single block or cell."""

    block_id: str
    instruction: str
    cell_id: str | None = None  # Optional: for editing table cells


@router.post("/{document_id}/ai-edit", response_model=DocumentJSON)
async def ai_edit_block(document_id: str, payload: AIEditRequest) -> DocumentJSON:
    """Apply an AI edit to a paragraph block or table cell.
    
    Uses LangGraph + Gemini as the primary AI backend.
    Evals are run automatically on every edit.
    """
    from services.document_edit_service import get_document_edit_service
    
    edit_service = get_document_edit_service()

    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")

        doc = DocumentJSON.model_validate_json(row.json)

        # Delegate to orchestration layer
        result = await edit_service.apply_ai_edit(
            doc=doc,
            block_id=payload.block_id,
            instruction=payload.instruction,
            cell_id=payload.cell_id,
        )
        
        if not result.success:
            if not result.validation_passed:
                raise HTTPException(
                    status_code=400,
                    detail={
                        "message": result.error,
                        "errors": result.validation_errors,
                    },
                )
            raise HTTPException(status_code=400, detail=result.error)

        # Persist updated document
        row.json = doc.model_dump_json()
        row.version = (row.version or 1) + 1

    return doc


class CheckboxUpdateRequest(BaseModel):
    """Request to update a checkbox state."""
    checkbox_id: str
    checked: bool


@router.post("/{document_id}/checkbox", response_model=DocumentJSON)
async def update_checkbox(document_id: str, payload: CheckboxUpdateRequest) -> DocumentJSON:
    """Update a checkbox state in the document."""
    
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")

        doc = DocumentJSON.model_validate_json(row.json)

        # Find and update the checkbox
        checkbox_found = False
        for checkbox in doc.checkboxes:
            if checkbox.id == payload.checkbox_id:
                checkbox.checked = payload.checked
                checkbox_found = True
                break

        if not checkbox_found:
            raise HTTPException(status_code=400, detail="Checkbox not found")

        row.json = doc.model_dump_json()
        row.version = (row.version or 1) + 1

    return doc


class DropdownUpdateRequest(BaseModel):
    """Request to update a dropdown selection."""
    dropdown_id: str
    selected: str


@router.post("/{document_id}/dropdown", response_model=DocumentJSON)
async def update_dropdown(document_id: str, payload: DropdownUpdateRequest) -> DocumentJSON:
    """Update a dropdown selection in the document."""
    
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")

        doc = DocumentJSON.model_validate_json(row.json)

        # Find and update the dropdown
        dropdown_found = False
        for dropdown in doc.dropdowns:
            if dropdown.id == payload.dropdown_id:
                # Validate that selected value is in options
                if payload.selected not in dropdown.options:
                    raise HTTPException(
                        status_code=400, 
                        detail=f"Invalid selection. Must be one of: {dropdown.options}"
                    )
                dropdown.selected = payload.selected
                dropdown_found = True
                break

        if not dropdown_found:
            raise HTTPException(status_code=400, detail="Dropdown not found")

        row.json = doc.model_dump_json()
        row.version = (row.version or 1) + 1

    return doc


class ValidationReportResponse(BaseModel):
    """Response containing detailed validation report."""
    document_id: str
    has_errors: bool
    has_warnings: bool
    stages: list[dict]
    issues: list[dict]


@router.get("/{document_id}/validate")
async def validate_document(document_id: str) -> ValidationReportResponse:
    """Validate the current document state against its original DOCX.
    
    This performs a full roundtrip validation:
    1. Compares stored JSON against original DOCX parsing
    2. Checks for any content loss or corruption
    3. Reports structural differences (paragraphs, tables, cells, etc.)
    
    Returns a detailed validation report with all issues found.
    """
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")
        
        # Get current JSON from DB
        current_json = DocumentJSON.model_validate_json(row.json)
        
        # Parse original DOCX fresh
        original_json = docx_to_json(row.base_docx_path, document_id)
        
        # Run full validation
        report = validate_parse_stage(row.base_docx_path, current_json)
        
        logger.info(f"[VALIDATE] Document {document_id}: errors={report.has_errors}, warnings={report.has_warnings}")
        
        return ValidationReportResponse(
            document_id=document_id,
            has_errors=report.has_errors,
            has_warnings=report.has_warnings,
            stages=[
                {
                    "stage": s.stage,
                    "total_chars": s.total_chars,
                    "paragraph_count": s.paragraph_count,
                    "table_count": s.table_count,
                    "row_count": s.row_count,
                    "cell_count": s.cell_count,
                    "run_count": s.run_count,
                    "checkbox_count": s.checkbox_count,
                    "dropdown_count": s.dropdown_count,
                }
                for s in report.stages
            ],
            issues=[
                {
                    "stage": i.stage,
                    "severity": i.severity,
                    "category": i.category,
                    "message": i.message,
                    "details": i.details,
                }
                for i in report.issues
            ],
        )


@router.post("/{document_id}/validate-export")
async def validate_export(document_id: str) -> ValidationReportResponse:
    """Perform a test export and validate the result without saving.
    
    This is useful for checking if an export will succeed before committing:
    1. Exports current JSON to a temporary DOCX
    2. Compares exported DOCX against original
    3. Reports any content loss or corruption
    
    Does NOT update the document version or save the export.
    """
    import tempfile
    import os
    
    with get_session() as db:
        row = db.get(Document, document_id)
        if row is None:
            raise HTTPException(status_code=404, detail="Document not found")
        
        current_json = DocumentJSON.model_validate_json(row.json)
        original_json = docx_to_json(row.base_docx_path, document_id)
        
        # Export to temp file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            temp_path = tmp.name
        
        try:
            apply_json_to_docx(current_json, row.base_docx_path, temp_path)
            
            # Validate export
            report = validate_export_stage(original_json, current_json, row.base_docx_path, temp_path)
            
            logger.info(f"[VALIDATE-EXPORT] Document {document_id}: errors={report.has_errors}, warnings={report.has_warnings}")
            
            return ValidationReportResponse(
                document_id=document_id,
                has_errors=report.has_errors,
                has_warnings=report.has_warnings,
                stages=[
                    {
                        "stage": s.stage,
                        "total_chars": s.total_chars,
                        "paragraph_count": s.paragraph_count,
                        "table_count": s.table_count,
                        "row_count": s.row_count,
                        "cell_count": s.cell_count,
                        "run_count": s.run_count,
                        "checkbox_count": s.checkbox_count,
                        "dropdown_count": s.dropdown_count,
                    }
                    for s in report.stages
                ],
                issues=[
                    {
                        "stage": i.stage,
                        "severity": i.severity,
                        "category": i.category,
                        "message": i.message,
                        "details": i.details,
                    }
                    for i in report.issues
                ],
            )
        finally:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.unlink(temp_path)
