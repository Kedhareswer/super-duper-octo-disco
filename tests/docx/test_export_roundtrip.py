import sys
import zipfile
from pathlib import Path

# Ensure project root (containing the `services` package) is on sys.path
ROOT = Path(__file__).resolve().parents[2]  # tests/docx/ -> tests/ -> project root
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.document_engine import docx_to_json, apply_json_to_docx


BASE_DIR = Path(__file__).resolve().parents[2]  # tests/docx/ -> tests/ -> project root
SAMPLE_DOCX = BASE_DIR / "data" / "uploads" / "docx" / "test2.docx"


def test_export_roundtrip_preserves_zip_structure(tmp_path):
    """Export should preserve DOCX zip structure. Only document.xml may differ."""

    assert SAMPLE_DOCX.exists(), f"Sample DOCX not found: {SAMPLE_DOCX}"

    # Parse once just to ensure docx_to_json does not throw
    doc = docx_to_json(str(SAMPLE_DOCX), document_id="test2.docx")
    assert doc.blocks, "Parsed document should have at least one block"

    out_path = tmp_path / "roundtrip.docx"
    result_path = Path(
        apply_json_to_docx(doc, base_docx_path=str(SAMPLE_DOCX), out_docx_path=str(out_path))
    )

    assert result_path.exists(), "Exported DOCX was not created"

    with zipfile.ZipFile(SAMPLE_DOCX, "r") as zin, zipfile.ZipFile(result_path, "r") as zout:
        in_names = sorted(i.filename for i in zin.infolist())
        out_names = sorted(i.filename for i in zout.infolist())
        assert in_names == out_names, "Zip entry lists differ between original and export"

        for name in in_names:
            in_bytes = zin.read(name)
            out_bytes = zout.read(name)
            if name != "word/document.xml":
                assert (
                    in_bytes == out_bytes
                ), f"Entry bytes differ for {name}, but only document.xml is allowed to change"


def test_parser_sees_final_funds_table():
    """Sanity check: parser should see at least one table with multiple rows for funds section."""

    doc = docx_to_json(str(SAMPLE_DOCX), document_id="test2.docx")
    tables = [b for b in doc.blocks if b.type == "table"]
    assert tables, "Expected at least one table in the sample DOCX"

    # Look for a table with >= 5 rows and >= 3 columns in the first row
    rich_tables = [
        t
        for t in tables
        if len(t.rows) >= 5 and t.rows and len(t.rows[0].cells) >= 3
    ]
    assert rich_tables, "Expected to find a larger funds-like table in the parsed JSON"


def test_edit_single_cell_and_export_contains_new_text(tmp_path):
    """Editing a single table cell in JSON should appear in exported document.xml."""

    assert SAMPLE_DOCX.exists(), f"Sample DOCX not found: {SAMPLE_DOCX}"

    doc = docx_to_json(str(SAMPLE_DOCX), document_id="test2.docx")

    # Find first table and first cell that has at least one paragraph
    tables = [b for b in doc.blocks if b.type == "table"]
    assert tables, "Expected at least one table in the sample DOCX"

    target_cell = None
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.blocks:
                    target_cell = cell
                    break
            if target_cell is not None:
                break
        if target_cell is not None:
            break

    assert target_cell is not None, "Could not find a table cell with content to edit"

    # Mutate JSON: change text of first paragraph's first run in the target cell
    para = target_cell.blocks[0]
    if para.runs:
        para.runs[0].text = "TEST_CELL_EDIT"
    else:
        from models.schemas import Run

        para.runs.append(
            Run(id="test-run", xml_ref="p[0]/r[0]", text="TEST_CELL_EDIT", bold=False, italic=False)
        )

    out_path = tmp_path / "edited_cell.docx"
    result_path = Path(
        apply_json_to_docx(doc, base_docx_path=str(SAMPLE_DOCX), out_docx_path=str(out_path))
    )
    assert result_path.exists(), "Exported DOCX was not created"

    # Ensure DOCX zip is still structurally valid and contains the new text in document.xml
    with zipfile.ZipFile(result_path, "r") as zout:
        names = [i.filename for i in zout.infolist()]
        assert "word/document.xml" in names, "Exported DOCX missing word/document.xml"
        xml_bytes = zout.read("word/document.xml")
        xml_text = xml_bytes.decode("utf-8", errors="ignore")
        assert "TEST_CELL_EDIT" in xml_text, "Edited cell text not found in exported document.xml"
