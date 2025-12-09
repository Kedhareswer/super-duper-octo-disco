"""Debug Excel XML structure - inspect raw XLSX XML content."""

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from io import BytesIO

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

TEST_FILE = Path("data/uploads/excel/excel_test.XLSX")


def print_element_tree(element, indent=0, max_depth=3):
    """Print element tree structure."""
    if indent > max_depth:
        return
    
    tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
    attrs = ' '.join(f'{k}="{v}"' for k, v in list(element.attrib.items())[:3])
    text = (element.text or '').strip()[:30]
    
    print(f"{'  ' * indent}<{tag} {attrs}>{text}")
    
    for child in list(element)[:5]:
        print_element_tree(child, indent + 1, max_depth)
    
    if len(list(element)) > 5:
        print(f"{'  ' * (indent + 1)}... ({len(list(element)) - 5} more children)")


def main():
    if not TEST_FILE.exists():
        print(f"Test file not found: {TEST_FILE}")
        return
    
    print(f"Analyzing: {TEST_FILE}")
    print("=" * 60)
    
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        print("\nZIP contents:")
        for info in zf.infolist():
            print(f"  {info.filename}: {info.file_size:,} bytes")
        
        # Analyze worksheet structure
        print("\n" + "=" * 60)
        print("Worksheet 1 structure:")
        print("=" * 60)
        
        try:
            content = zf.read("xl/worksheets/sheet1.xml")
            tree = ET.parse(BytesIO(content))
            root = tree.getroot()
            print_element_tree(root, max_depth=4)
        except Exception as e:
            print(f"Error: {e}")
        
        # Show sheetData sample
        print("\n" + "=" * 60)
        print("Sample cell data (first 500 chars of sheetData):")
        print("=" * 60)
        
        try:
            content = zf.read("xl/worksheets/sheet1.xml").decode('utf-8')
            start = content.find('<sheetData')
            if start > 0:
                print(content[start:start+500])
        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    main()
