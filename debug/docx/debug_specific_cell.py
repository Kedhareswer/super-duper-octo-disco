"""Debug specific cell that's losing text."""

import sys
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET
from io import BytesIO

sys.path.insert(0, str(Path(__file__).parent))

from services.document_engine import docx_to_json, NS, _find_node_by_ref
from models.schemas import TableBlock

TEST_FILE = Path("data/uploads/docx/test2.docx")


def main():
    # Parse document
    doc = docx_to_json(str(TEST_FILE), "test")
    
    # Find block[4] which is a table
    block = doc.blocks[4]
    print(f"Block 4 type: {type(block).__name__}")
    
    if isinstance(block, TableBlock):
        print(f"Table has {len(block.rows)} rows")
        
        # Look at row 4, cell 2
        if len(block.rows) > 4:
            row = block.rows[4]
            print(f"\nRow 4 has {len(row.cells)} cells")
            
            if len(row.cells) > 2:
                cell = row.cells[2]
                print(f"\nCell 2:")
                print(f"  xml_ref: {cell.xml_ref}")
                print(f"  blocks: {len(cell.blocks)}")
                
                for i, para in enumerate(cell.blocks):
                    print(f"\n  Para {i}:")
                    print(f"    xml_ref: {para.xml_ref}")
                    print(f"    runs: {len(para.runs)}")
                    for j, run in enumerate(para.runs):
                        print(f"      Run {j}: xml_ref={run.xml_ref}, text='{run.text}'")
    
    # Now look at the actual XML
    print("\n" + "="*60)
    print("Looking at actual XML structure")
    print("="*60)
    
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        with zf.open('word/document.xml') as f:
            tree = ET.parse(f)
    
    root = tree.getroot()
    body = root.find("w:body", NS)
    
    # Find the cell
    cell_ref = "tbl[1]/tr[4]/tc[2]"  # block[4] is tbl[1] (second table)
    print(f"\nLooking for: {cell_ref}")
    
    # Actually, let's trace through the tables
    tables = body.findall("w:tbl", NS)
    print(f"Found {len(tables)} tables in body")
    
    # block[4] means it's the 2nd table (after some paragraphs)
    # Let's check the xml_ref from the parsed data
    if isinstance(doc.blocks[4], TableBlock):
        tbl_ref = doc.blocks[4].xml_ref
        print(f"Table xml_ref: {tbl_ref}")
        
        # Find the table
        tbl_el = _find_node_by_ref(body, tbl_ref)
        if tbl_el is not None:
            print("Found table element")
            
            rows = tbl_el.findall("w:tr", NS)
            print(f"Table has {len(rows)} rows")
            
            if len(rows) > 4:
                row_el = rows[4]
                cells = row_el.findall(".//w:tc", NS)
                print(f"Row 4 has {len(cells)} cells")
                
                if len(cells) > 2:
                    cell_el = cells[2]
                    paras = cell_el.findall(".//w:p", NS)
                    print(f"Cell 2 has {len(paras)} paragraphs")
                    
                    for pi, p_el in enumerate(paras):
                        runs = p_el.findall(".//w:r", NS)
                        print(f"\n  Para {pi} has {len(runs)} runs")
                        for ri, r_el in enumerate(runs):
                            t_els = r_el.findall("w:t", NS)
                            text = "".join(t.text or "" for t in t_els)
                            print(f"    Run {ri}: '{text}'")


if __name__ == "__main__":
    main()
