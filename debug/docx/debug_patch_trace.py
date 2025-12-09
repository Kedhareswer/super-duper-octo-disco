"""Trace the patching process."""

import sys
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET
from io import BytesIO

sys.path.insert(0, str(Path('.')))
from services.document_engine import docx_to_json, NS, _find_node_by_ref
from models.schemas import TableBlock, ParagraphBlock

TEST_FILE = Path("data/uploads/docx/test2.docx")


def main():
    # Parse document
    doc = docx_to_json(str(TEST_FILE), "test")
    
    # Load XML
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)
    
    root = tree.getroot()
    body = root.find("w:body", NS)
    
    # Find the specific cell
    block = doc.blocks[4]  # tbl[1]
    if not isinstance(block, TableBlock):
        print("Block 4 is not a table!")
        return
    
    cell = block.rows[4].cells[2]
    para = cell.blocks[0]
    
    print(f"Cell xml_ref: {cell.xml_ref}")
    print(f"Para xml_ref: {para.xml_ref}")
    print(f"Para runs: {len(para.runs)}")
    for i, run in enumerate(para.runs):
        print(f"  Run {i}: xml_ref={run.xml_ref}, text='{run.text}'")
    
    # Try to find the paragraph using _find_node_by_ref
    print(f"\nTrying to find: {para.xml_ref}")
    p_el = _find_node_by_ref(body, para.xml_ref)
    
    if p_el is None:
        print("  FAILED to find paragraph!")
        
        # Try to find the cell first
        print(f"\nTrying to find cell: {cell.xml_ref}")
        tc_el = _find_node_by_ref(body, cell.xml_ref)
        if tc_el is None:
            print("  FAILED to find cell!")
        else:
            print("  Found cell!")
            # Look for paragraphs in cell
            ps = tc_el.findall(".//w:p", NS)
            print(f"  Cell has {len(ps)} paragraphs")
            for i, p in enumerate(ps):
                rs = p.findall(".//w:r", NS)
                print(f"    Para {i}: {len(rs)} runs")
                for j, r in enumerate(rs):
                    ts = r.findall("w:t", NS)
                    text = "".join(t.text or "" for t in ts)
                    print(f"      Run {j}: '{text}'")
    else:
        print("  Found paragraph!")
        rs = p_el.findall(".//w:r", NS)
        print(f"  Has {len(rs)} runs")


if __name__ == "__main__":
    main()
