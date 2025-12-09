"""Debug Excel XML namespaces - check what namespaces are in XLSX files."""

import sys
import zipfile
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

TEST_FILE = Path("data/uploads/excel/excel_test.XLSX")


def extract_namespaces(xml_content: str) -> dict:
    """Extract all xmlns declarations from XML content."""
    ns_pattern = r'xmlns:?(\w*)="([^"]+)"'
    matches = re.findall(ns_pattern, xml_content[:5000])
    return {prefix or "default": uri for prefix, uri in matches}


def main():
    if not TEST_FILE.exists():
        print(f"Test file not found: {TEST_FILE}")
        return
    
    print(f"Analyzing: {TEST_FILE}")
    print("=" * 60)
    
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        # Check key XML files
        xml_files = [
            "xl/workbook.xml",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet3.xml",
            "xl/sharedStrings.xml",
            "xl/styles.xml",
        ]
        
        for xml_file in xml_files:
            try:
                content = zf.read(xml_file).decode('utf-8')
                namespaces = extract_namespaces(content)
                
                print(f"\n{xml_file}:")
                for prefix, uri in sorted(namespaces.items()):
                    print(f"  xmlns:{prefix} = {uri}")
            except KeyError:
                print(f"\n{xml_file}: NOT FOUND")


if __name__ == "__main__":
    main()
