"""Debug dropdown XML structure."""

import sys
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path('.')))
from services.document_engine import NS

TEST_FILE = Path("data/uploads/docx/test2.docx")


def main():
    with zipfile.ZipFile(TEST_FILE, 'r') as zf:
        with zf.open('word/document.xml') as f:
            content = f.read()
    
    # Find all SDT elements with dropdowns
    tree = ET.parse(zipfile.ZipFile(TEST_FILE).open('word/document.xml'))
    root = tree.getroot()
    
    print("Looking for dropdown SDT elements...")
    
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is None:
            continue
        
        combo = sdt_pr.find("w:comboBox", NS)
        dropdown = sdt_pr.find("w:dropDownList", NS)
        
        if combo is None and dropdown is None:
            continue
        
        id_el = sdt_pr.find("w:id", NS)
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "") if id_el else "?"
        
        print(f"\n{'='*60}")
        print(f"Dropdown SDT id={field_id}")
        print(f"{'='*60}")
        
        # Show the structure
        sdt_content = sdt.find("w:sdtContent", NS)
        if sdt_content is not None:
            print("\nsdtContent structure:")
            
            # Check for table cells
            tcs = sdt_content.findall(".//w:tc", NS)
            print(f"  Contains {len(tcs)} table cells")
            
            # Check for paragraphs
            ps = sdt_content.findall(".//w:p", NS)
            print(f"  Contains {len(ps)} paragraphs")
            
            # Check for runs
            rs = sdt_content.findall(".//w:r", NS)
            print(f"  Contains {len(rs)} runs")
            
            # Check for text
            ts = sdt_content.findall(".//w:t", NS)
            print(f"  Contains {len(ts)} text elements")
            
            # Show all text content
            print("\n  Text content:")
            for i, t in enumerate(ts):
                print(f"    [{i}] '{t.text}'")
            
            # Show the XML snippet
            print("\n  Raw XML (first 2000 chars):")
            xml_str = ET.tostring(sdt_content, encoding='unicode')
            print(xml_str[:2000])


if __name__ == "__main__":
    main()
