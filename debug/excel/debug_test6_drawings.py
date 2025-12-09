"""Debug test6.xlsx drawing/image parsing issue."""

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from io import BytesIO

ROOT = Path(__file__).parent.parent.parent
TEST_FILE = ROOT / "data" / "uploads" / "excel" / "test6.xlsx"

def debug_drawings():
    """Debug drawing files and relationships."""
    print(f"Debugging drawings for: {TEST_FILE}")
    
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        # List all rels files
        rels_files = [f for f in zf.namelist() if '.rels' in f]
        print(f"\nRels files: {len(rels_files)}")
        for rf in rels_files:
            print(f"  - {rf}")
        
        # Check worksheet rels to find drawing references
        print("\n--- Checking worksheet rels for drawings ---")
        for sheet_num in range(1, 16):
            sheet_rels = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
            if sheet_rels in zf.namelist():
                content = zf.read(sheet_rels).decode('utf-8')
                if 'drawing' in content.lower():
                    print(f"\nSheet {sheet_num} has drawing reference:")
                    # Parse the rels file
                    try:
                        tree = ET.parse(BytesIO(content.encode()))
                        for rel in tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if 'drawing' in rel.get('Target', '').lower():
                                print(f"  Target: {rel.get('Target')}")
                    except:
                        print(f"  (Could not parse rels)")
        
        # Check drawing rels files
        print("\n--- Checking drawing rels files ---")
        drawing_rels = [f for f in zf.namelist() if 'drawings/_rels' in f]
        for dr in drawing_rels:
            print(f"\n{dr}:")
            try:
                content = zf.read(dr).decode('utf-8')
                # Parse and show relationships
                tree = ET.parse(BytesIO(content.encode()))
                for rel in tree.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    print(f"  Type: {rel.get('Type', 'N/A').split('/')[-1]}")
                    print(f"  Target: {rel.get('Target')}")
            except Exception as e:
                print(f"  Error: {e}")
        
        # Try to parse each drawing rels as raw XML
        print("\n--- Checking vmlDrawing files ---")
        vml_files = [f for f in zf.namelist() if 'vmlDrawing' in f]
        for vml in vml_files:
            print(f"\n{vml}:")
            try:
                content = zf.read(vml)
                ET.parse(BytesIO(content))
                print(f"  ✓ Valid XML")
            except ET.ParseError as e:
                print(f"  ✗ ParseError: {e}")
                # Show problematic lines
                try:
                    text = content.decode('utf-8', errors='replace')
                    lines = text.split('\n')
                    print(f"  Lines around error:")
                    for i, line in enumerate(lines[15:25], start=16):
                        prefix = ">>>" if i == 20 else "   "
                        print(f"    {prefix} {i}: {line[:150]}")
                except:
                    pass


if __name__ == "__main__":
    debug_drawings()
