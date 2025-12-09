"""Debug what happens to the cell after export."""

import sys
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET
from io import BytesIO

sys.path.insert(0, str(Path(__file__).parent))

from services.document_engine import docx_to_json, apply_json_to_docx, NS, _find_node_by_ref
from models.schemas import TableBlock

TEST_FILE = Path("data/uploads/docx/test2.docx")
EXPORT_FILE = Path("data/test_outputs/debug_export.docx")


def check_cell_in_file(file_path, label):
    """Check the specific cell content in a file."""
    print(f"\n{'='*60}")
    print(f"  {label}: {file_path}")
    print(f"{'='*60}")
    
    with zipfile.ZipFile(file_path, 'r') as zf:
        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)
    
    root = tree.getroot()
    body = root.find("w:body", NS)
    
    # Find tbl[1]/tr[4]/tc[2]
    tables = body.findall("w:tbl", NS)
    if len(tables) < 2:
        print("  Not enough tables!")
        return
    
    tbl = tables[1]  # tbl[1]
    rows = tbl.findall("w:tr", NS)
    if len(rows) < 5:
        print("  Not enough rows!")
        return
    
    row = rows[4]  # tr[4]
    cells = row.findall(".//w:tc", NS)
    if len(cells) < 3:
        print("  Not enough cells!")
        return
    
    cell = cells[2]  # tc[2]
    
    # Get all paragraphs
    paras = cell.findall(".//w:p", NS)
    print(f"  Cell has {len(paras)} paragraphs")
    
    for pi, p in enumerate(paras):
        runs = p.findall(".//w:r", NS)
        print(f"\n  Para {pi}: {len(runs)} runs")
        for ri, r in enumerate(runs):
            t_els = r.findall("w:t", NS)
            text = "".join(t.text or "" for t in t_els)
            print(f"    Run {ri}: '{text}'")
        
        # Also check for any w:t directly in paragraph (unusual but possible)
        direct_t = p.findall("w:t", NS)
        if direct_t:
            print(f"    Direct w:t in para: {[t.text for t in direct_t]}")


def main():
    # Check original
    check_cell_in_file(TEST_FILE, "ORIGINAL")
    
    # Parse and export
    doc = docx_to_json(str(TEST_FILE), "test")
    
    # Show what we parsed for this cell
    print(f"\n{'='*60}")
    print("  PARSED JSON for tbl[1]/tr[4]/tc[2]")
    print(f"{'='*60}")
    
    block = doc.blocks[4]
    if isinstance(block, TableBlock):
        cell = block.rows[4].cells[2]
        print(f"  Cell xml_ref: {cell.xml_ref}")
        for pi, para in enumerate(cell.blocks):
            print(f"\n  Para {pi}: xml_ref={para.xml_ref}")
            for ri, run in enumerate(para.runs):
                print(f"    Run {ri}: xml_ref={run.xml_ref}, text='{run.text}'")
    
    # Export
    EXPORT_FILE.parent.mkdir(parents=True, exist_ok=True)
    apply_json_to_docx(doc, str(TEST_FILE), str(EXPORT_FILE))
    
    # Check export
    check_cell_in_file(EXPORT_FILE, "EXPORTED")


if __name__ == "__main__":
    main()
