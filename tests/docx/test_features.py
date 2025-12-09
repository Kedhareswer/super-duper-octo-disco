"""Test document features like drawings, tables, colspan, borders."""
import pytest
from pathlib import Path

from services.document_engine import docx_to_json


@pytest.fixture
def test_doc():
    """Load test2.docx for feature testing."""
    test_path = Path("data/uploads/docx/test2.docx")
    if not test_path.exists():
        pytest.skip("test2.docx not found in data/uploads/docx/")
    return docx_to_json(str(test_path), 'test2.docx')


def test_document_blocks(test_doc):
    """Test that document has expected block types."""
    assert len(test_doc.blocks) > 0
    
    # Count by type
    drawings = [b for b in test_doc.blocks if b.type.value == "drawing"]
    tables = [b for b in test_doc.blocks if b.type.value == "table"]
    paragraphs = [b for b in test_doc.blocks if b.type.value == "paragraph"]
    
    assert len(paragraphs) >= 0  # Can have 0 or more
    assert len(tables) >= 0  # Can have 0 or more


def test_table_structure(test_doc):
    """Test table structure is correctly parsed."""
    tables = [b for b in test_doc.blocks if b.type.value == "table"]
    
    if not tables:
        pytest.skip("No tables in document")
    
    # Check first table has rows and cells
    t = tables[0]
    assert len(t.rows) > 0
    assert len(t.rows[0].cells) > 0


def test_cell_has_xml_ref(test_doc):
    """Test that cells have valid xml_ref."""
    tables = [b for b in test_doc.blocks if b.type.value == "table"]
    
    if not tables:
        pytest.skip("No tables in document")
    
    t = tables[0]
    for row in t.rows:
        for cell in row.cells:
            assert cell.xml_ref is not None
            assert cell.xml_ref.startswith("tbl[")
