"""Debug test6.xlsx parsing issue."""

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from io import BytesIO

ROOT = Path(__file__).parent.parent.parent
TEST_FILE = ROOT / "data" / "uploads" / "excel" / "test6.xlsx"

def debug_xlsx(file_path: Path):
    """Debug an XLSX file that fails to parse."""
    print(f"Debugging: {file_path}")
    print(f"File size: {file_path.stat().st_size:,} bytes")
    
    with zipfile.ZipFile(file_path, 'r') as zf:
        print(f"\nFiles in archive: {len(zf.namelist())}")
        
        # List XML files
        xml_files = [f for f in zf.namelist() if f.endswith('.xml')]
        print(f"XML files: {len(xml_files)}")
        
        # Try to parse each XML file
        print("\nParsing each XML file...")
        for xml_file in xml_files:
            try:
                content = zf.read(xml_file)
                ET.parse(BytesIO(content))
                print(f"  ✓ {xml_file}")
            except ET.ParseError as e:
                print(f"  ✗ {xml_file}: {e}")
                # Show the problematic area
                try:
                    text = content.decode('utf-8', errors='replace')
                    lines = text.split('\n')
                    if hasattr(e, 'position'):
                        line_no, col = e.position
                        if line_no <= len(lines):
                            print(f"\n    Context around line {line_no}:")
                            start = max(0, line_no - 3)
                            end = min(len(lines), line_no + 2)
                            for i in range(start, end):
                                prefix = ">>>" if i + 1 == line_no else "   "
                                print(f"    {prefix} {i+1}: {lines[i][:200]}")
                except:
                    pass
            except Exception as e:
                print(f"  ✗ {xml_file}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    if TEST_FILE.exists():
        debug_xlsx(TEST_FILE)
    else:
        print(f"File not found: {TEST_FILE}")
