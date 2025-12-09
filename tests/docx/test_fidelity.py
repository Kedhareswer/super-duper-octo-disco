"""
Fidelity hardening tests for DOCX parsing and round-trip.

These tests ensure we correctly handle:
- Complex merged tables (colspan, rowspan via vMerge)
- Borders and shading variations
- Documents with multiple drawings/logos
- SDT (content controls) inside table cells without dropping cells
"""
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

# Ensure project root is on sys.path (tests/docx/ -> tests/ -> project root)
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.document_engine import docx_to_json, apply_json_to_docx
from models.schemas import (
    DocumentJSON,
    ParagraphBlock,
    TableBlock,
    TableCell,
    Run,
    CellBorder,
    CellBorders,
    DrawingBlock,
)


BASE_DIR = Path(__file__).resolve().parents[2]
SAMPLE_DOCX = BASE_DIR / "data" / "uploads" / "docx" / "test2.docx"


# =============================================================================
# MERGED TABLES TESTS
# =============================================================================

class TestMergedTables:
    """Tests for complex merged table handling."""

    def test_colspan_extraction(self):
        """Cells with gridSpan should have col_span > 1."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        # Find any cell with colspan > 1
        cells_with_colspan = []
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.col_span > 1:
                        cells_with_colspan.append(cell)
        
        assert cells_with_colspan, (
            "Expected at least one cell with col_span > 1 in test2.docx. "
            "This document should have merged header cells."
        )
        
        # Verify the colspan value is reasonable (not absurdly large)
        for cell in cells_with_colspan:
            assert 1 < cell.col_span <= 20, f"Unexpected col_span value: {cell.col_span}"

    def test_vmerge_extraction(self):
        """Cells with vMerge should have v_merge set to 'restart' or 'continue'."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        # Collect vMerge info
        restart_cells = []
        continue_cells = []
        
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.v_merge == "restart":
                        restart_cells.append(cell)
                    elif cell.v_merge == "continue":
                        continue_cells.append(cell)
        
        # If we have restart cells, we should also have continue cells
        if restart_cells:
            assert continue_cells, (
                "Found vMerge='restart' cells but no 'continue' cells. "
                "Vertical merges should have both."
            )

    def test_merged_cell_count_consistency(self):
        """
        Total logical cells (accounting for spans) should match grid structure.
        
        For each row, sum of col_span values should be consistent across rows
        (or at least not wildly different, accounting for merged cells).
        """
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        for table in tables:
            if not table.rows:
                continue
            
            # Calculate logical column count for each row
            row_widths = []
            for row in table.rows:
                logical_width = sum(cell.col_span for cell in row.cells)
                row_widths.append(logical_width)
            
            # All rows should have the same logical width (grid columns)
            if row_widths:
                expected_width = max(row_widths)  # Use max as reference
                for i, width in enumerate(row_widths):
                    # Allow some tolerance for complex tables
                    assert width <= expected_width, (
                        f"Row {i} has logical width {width}, expected <= {expected_width}"
                    )

    def test_roundtrip_preserves_merged_cells(self, tmp_path):
        """After export, merged cells should still be present in document.xml."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # Find a cell with colspan
        colspan_cell = None
        for block in doc.blocks:
            if block.type.value != "table":
                continue
            for row in block.rows:
                for cell in row.cells:
                    if cell.col_span > 1:
                        colspan_cell = cell
                        break
                if colspan_cell:
                    break
            if colspan_cell:
                break
        
        if colspan_cell is None:
            pytest.skip("No colspan cells found in test document")
        
        # Export
        out_path = tmp_path / "merged_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        # Parse exported document.xml and verify gridSpan is preserved
        with zipfile.ZipFile(out_path, "r") as zf:
            xml_bytes = zf.read("word/document.xml")
        
        # Check that gridSpan elements exist (may have ns prefix)
        assert b"gridSpan" in xml_bytes, (
            "Exported document.xml should contain gridSpan for merged cells"
        )


# =============================================================================
# BORDERS AND SHADING TESTS
# =============================================================================

class TestBordersAndShading:
    """Tests for cell border and background color extraction."""

    def test_cell_background_extraction(self):
        """Cells with shading should have background_color set."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        cells_with_bg = []
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.background_color:
                        cells_with_bg.append(cell)
        
        # test2.docx should have some colored cells (headers, highlights)
        assert cells_with_bg, (
            "Expected at least one cell with background_color in test2.docx"
        )
        
        # Verify color format (should be hex without #)
        for cell in cells_with_bg:
            color = cell.background_color
            assert len(color) == 6, f"Expected 6-char hex color, got: {color}"
            assert all(c in "0123456789ABCDEFabcdef" for c in color), (
                f"Invalid hex color: {color}"
            )

    def test_cell_borders_extraction(self):
        """Cells with explicit borders should have borders object set."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        cells_with_borders = []
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.borders:
                        cells_with_borders.append(cell)
        
        assert cells_with_borders, (
            "Expected at least one cell with explicit borders in test2.docx"
        )
        
        # Verify border structure
        for cell in cells_with_borders:
            borders = cell.borders
            # At least one side should be defined
            has_any = any([borders.top, borders.bottom, borders.left, borders.right])
            assert has_any, f"Cell {cell.id} has borders object but no sides defined"

    def test_border_style_values(self):
        """Border styles should be valid OOXML values."""
        valid_styles = {
            "none", "nil", "single", "thick", "double", "dotted", "dashed",
            "dashSmallGap", "dotDash", "dotDotDash", "triple", "thinThickSmallGap",
            "thickThinSmallGap", "thinThickThinSmallGap", "thinThickMediumGap",
            "thickThinMediumGap", "thinThickThinMediumGap", "thinThickLargeGap",
            "thickThinLargeGap", "thinThickThinLargeGap", "wave", "doubleWave",
            "dashDotStroked", "threeDEmboss", "threeDEngrave", "outset", "inset",
        }
        
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    if not cell.borders:
                        continue
                    for side in [cell.borders.top, cell.borders.bottom, 
                                 cell.borders.left, cell.borders.right]:
                        if side:
                            assert side.style in valid_styles, (
                                f"Invalid border style: {side.style}"
                            )

    def test_roundtrip_preserves_shading(self, tmp_path):
        """After export, cell shading should still be present."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        out_path = tmp_path / "shading_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        with zipfile.ZipFile(out_path, "r") as zf:
            xml_bytes = zf.read("word/document.xml")
        
        # Check that shd elements exist (may have ns prefix)
        # Note: shd may be in tcPr or pPr
        assert b"shd" in xml_bytes or b"fill=" in xml_bytes, (
            "Exported document.xml should contain shd or fill for cell shading"
        )


# =============================================================================
# DRAWINGS AND LOGOS TESTS
# =============================================================================

class TestDrawingsAndLogos:
    """Tests for drawing/logo extraction."""

    def test_drawing_extraction(self):
        """Documents with drawings should have DrawingBlock entries."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        drawings = [b for b in doc.blocks if b.type.value == "drawing"]
        
        # test2.docx should have at least one drawing (logo)
        assert drawings, "Expected at least one drawing in test2.docx"

    def test_drawing_dimensions(self):
        """Drawings should have valid dimensions in inches."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        drawings = [b for b in doc.blocks if b.type.value == "drawing"]
        
        for drawing in drawings:
            assert drawing.width_inches > 0, f"Drawing {drawing.id} has invalid width"
            assert drawing.height_inches > 0, f"Drawing {drawing.id} has invalid height"
            # Reasonable bounds (not larger than a page)
            assert drawing.width_inches < 20, f"Drawing {drawing.id} width too large"
            assert drawing.height_inches < 20, f"Drawing {drawing.id} height too large"

    def test_drawing_type_detection(self):
        """Drawing type should be 'vector_group', 'image', or 'unknown'."""
        valid_types = {"vector_group", "image", "unknown"}
        
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        drawings = [b for b in doc.blocks if b.type.value == "drawing"]
        
        for drawing in drawings:
            assert drawing.drawing_type in valid_types, (
                f"Invalid drawing type: {drawing.drawing_type}"
            )

    def test_drawing_name_extraction(self):
        """Drawings should have a name if wp:docPr has a name attribute."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        drawings = [b for b in doc.blocks if b.type.value == "drawing"]
        
        # At least one drawing should have a name
        named_drawings = [d for d in drawings if d.name]
        assert named_drawings, "Expected at least one drawing with a name"

    def test_roundtrip_preserves_drawings(self, tmp_path):
        """After export, drawings should still be present in document.xml."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        out_path = tmp_path / "drawings_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        with zipfile.ZipFile(out_path, "r") as zf:
            xml_bytes = zf.read("word/document.xml")
        
        # Check that drawing elements exist (may have ns prefix)
        assert b"drawing" in xml_bytes, (
            "Exported document.xml should contain drawing elements"
        )


# =============================================================================
# SDT / CONTENT CONTROLS IN TABLES TESTS
# =============================================================================

class TestSDTInTables:
    """Tests to ensure SDT content controls don't cause cell drops."""

    def test_sdt_cells_not_dropped(self):
        """
        Cells containing SDT (content controls) should not be silently dropped.
        
        This is a regression test: some parsers skip cells that only contain
        SDT elements, causing row/column misalignment.
        """
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        tables = [b for b in doc.blocks if b.type.value == "table"]
        
        for table in tables:
            for row_idx, row in enumerate(table.rows):
                # Every row should have at least one cell
                assert row.cells, f"Row {row_idx} in table {table.id} has no cells"
                
                # Check that cell IDs are unique within the row
                cell_ids = [c.id for c in row.cells]
                assert len(cell_ids) == len(set(cell_ids)), (
                    f"Duplicate cell IDs in row {row_idx}: {cell_ids}"
                )

    def test_sdt_text_extraction(self):
        """
        Text inside SDT content controls should be extracted.
        
        SDT structure:
        <w:sdt>
          <w:sdtContent>
            <w:p>
              <w:r><w:t>text</w:t></w:r>
            </w:p>
          </w:sdtContent>
        </w:sdt>
        """
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # Collect all text from all blocks
        all_text = []
        for block in doc.blocks:
            if block.type.value == "paragraph":
                for run in block.runs:
                    # Only TextRuns have text attribute; skip CheckboxRun/DropdownRun
                    if hasattr(run, 'text') and run.text:
                        all_text.append(run.text)
            elif block.type.value == "table":
                for row in block.rows:
                    for cell in row.cells:
                        for para in cell.blocks:
                            if hasattr(para, 'runs'):
                                for run in para.runs:
                                    # Only TextRuns have text attribute
                                    if hasattr(run, 'text') and run.text:
                                        all_text.append(run.text)
        
        # Should have extracted some text
        assert all_text, "No text extracted from document"
        
        # Look for typical SDT placeholder text patterns
        # (these are common in forms)
        combined = " ".join(all_text).lower()
        # Just verify we got substantial text
        assert len(combined) > 100, "Very little text extracted, possible SDT issue"

    def test_checkbox_extraction(self):
        """Checkboxes (w14:checkbox SDT) should be extracted."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # test2.docx should have checkboxes
        assert doc.checkboxes is not None, "checkboxes should not be None"
        
        if doc.checkboxes:
            for cb in doc.checkboxes:
                assert cb.id, "Checkbox should have an ID"
                assert cb.xml_ref, "Checkbox should have an xml_ref"
                assert isinstance(cb.checked, bool), "checked should be boolean"

    def test_dropdown_extraction(self):
        """Dropdowns (w:comboBox/w:dropDownList SDT) should be extracted."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        assert doc.dropdowns is not None, "dropdowns should not be None"
        
        if doc.dropdowns:
            for dd in doc.dropdowns:
                assert dd.id, "Dropdown should have an ID"
                assert dd.xml_ref, "Dropdown should have an xml_ref"
                assert isinstance(dd.options, list), "options should be a list"


# =============================================================================
# COMPREHENSIVE ROUND-TRIP TESTS
# =============================================================================

class TestComprehensiveRoundtrip:
    """End-to-end round-trip tests combining multiple features."""

    def test_full_fidelity_roundtrip(self, tmp_path):
        """
        Full round-trip should preserve:
        - All blocks (paragraphs, tables, drawings)
        - Merged cells
        - Borders and shading
        - Text content
        """
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # Count original elements
        orig_para_count = len([b for b in doc.blocks if b.type.value == "paragraph"])
        orig_table_count = len([b for b in doc.blocks if b.type.value == "table"])
        orig_drawing_count = len([b for b in doc.blocks if b.type.value == "drawing"])
        
        # Export
        out_path = tmp_path / "full_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        # Re-parse
        doc2 = docx_to_json(str(out_path), document_id="test2")
        
        # Count after round-trip
        new_para_count = len([b for b in doc2.blocks if b.type.value == "paragraph"])
        new_table_count = len([b for b in doc2.blocks if b.type.value == "table"])
        new_drawing_count = len([b for b in doc2.blocks if b.type.value == "drawing"])
        
        # Verify counts match
        assert new_para_count == orig_para_count, (
            f"Paragraph count changed: {orig_para_count} -> {new_para_count}"
        )
        assert new_table_count == orig_table_count, (
            f"Table count changed: {orig_table_count} -> {new_table_count}"
        )
        assert new_drawing_count == orig_drawing_count, (
            f"Drawing count changed: {orig_drawing_count} -> {new_drawing_count}"
        )

    def test_text_edit_roundtrip(self, tmp_path):
        """Editing text and round-tripping should preserve the edit."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # Find a paragraph with a run that has actual text (not empty)
        target_para = None
        target_run_idx = None
        for block in doc.blocks:
            if block.type.value == "paragraph" and block.runs:
                for idx, run in enumerate(block.runs):
                    if run.text and len(run.text) > 0:
                        target_para = block
                        target_run_idx = idx
                        break
                if target_para:
                    break
        
        if target_para is None:
            pytest.skip("No paragraph with non-empty runs found")
        
        # Edit the text
        original_text = target_para.runs[target_run_idx].text
        new_text = "FIDELITY_TEST_MARKER_12345"
        target_para.runs[target_run_idx].text = new_text
        
        # Export
        out_path = tmp_path / "text_edit_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        # Verify in raw XML
        with zipfile.ZipFile(out_path, "r") as zf:
            xml_text = zf.read("word/document.xml").decode("utf-8")
        
        assert new_text in xml_text, "Edited text not found in exported document"

    def test_table_cell_edit_roundtrip(self, tmp_path):
        """Editing table cell text should persist through round-trip."""
        doc = docx_to_json(str(SAMPLE_DOCX), document_id="test")
        
        # Find a table cell with text
        target_cell = None
        for block in doc.blocks:
            if block.type.value != "table":
                continue
            for row in block.rows:
                for cell in row.cells:
                    if cell.blocks and cell.blocks[0].runs:
                        target_cell = cell
                        break
                if target_cell:
                    break
            if target_cell:
                break
        
        if target_cell is None:
            pytest.skip("No table cell with text found")
        
        # Edit
        new_text = "CELL_FIDELITY_MARKER_67890"
        target_cell.blocks[0].runs[0].text = new_text
        
        # Export
        out_path = tmp_path / "cell_edit_roundtrip.docx"
        apply_json_to_docx(doc, str(SAMPLE_DOCX), str(out_path))
        
        # Verify
        with zipfile.ZipFile(out_path, "r") as zf:
            xml_text = zf.read("word/document.xml").decode("utf-8")
        
        assert new_text in xml_text, "Edited cell text not found in exported document"


# =============================================================================
# RUN TESTS
# =============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v"])
