"""Debug script to test the document pipeline at each stage."""

import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from io import BytesIO

# Test with the uploaded file
UPLOAD_PATH = Path("data/uploads/docx/test2.docx")
EXPORT_PATH = Path("data/exports/test2.docx.v9.docx")

def check_docx_validity(path: Path, label: str):
    """Check if a DOCX file is valid and can be opened."""
    print(f"\n{'='*60}")
    print(f"Checking: {label} ({path})")
    print(f"{'='*60}")
    
    if not path.exists():
        print(f"  ERROR: File does not exist!")
        return False
    
    print(f"  File size: {path.stat().st_size:,} bytes")
    
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            # Check for required files
            required = ['word/document.xml', '[Content_Types].xml']
            for req in required:
                if req in zf.namelist():
                    print(f"  ✓ {req} present")
                else:
                    print(f"  ✗ {req} MISSING!")
                    return False
            
            # List all files
            print(f"\n  Files in archive ({len(zf.namelist())} total):")
            for name in sorted(zf.namelist())[:20]:
                info = zf.getinfo(name)
                print(f"    - {name} ({info.file_size:,} bytes)")
            if len(zf.namelist()) > 20:
                print(f"    ... and {len(zf.namelist()) - 20} more files")
            
            # Parse document.xml
            print(f"\n  Parsing word/document.xml...")
            with zf.open('word/document.xml') as doc_xml:
                content = doc_xml.read()
                print(f"    Size: {len(content):,} bytes")
                
                # Check for XML declaration
                if content.startswith(b'<?xml'):
                    print(f"    ✓ Has XML declaration")
                else:
                    print(f"    ✗ Missing XML declaration (starts with: {content[:50]})")
                
                # Try to parse
                try:
                    tree = ET.parse(BytesIO(content))
                    root = tree.getroot()
                    print(f"    ✓ Valid XML")
                    print(f"    Root tag: {root.tag}")
                    
                    # Check namespaces
                    print(f"\n    Namespaces in document:")
                    for prefix, uri in root.attrib.items():
                        if prefix.startswith('{'):
                            continue
                        print(f"      {prefix}: {uri[:50]}...")
                    
                except ET.ParseError as e:
                    print(f"    ✗ XML Parse Error: {e}")
                    # Show problematic area
                    print(f"    First 500 chars: {content[:500]}")
                    return False
            
            # Check Content_Types
            print(f"\n  Checking [Content_Types].xml...")
            with zf.open('[Content_Types].xml') as ct_xml:
                ct_content = ct_xml.read()
                try:
                    ct_tree = ET.parse(BytesIO(ct_content))
                    print(f"    ✓ Valid XML")
                except ET.ParseError as e:
                    print(f"    ✗ XML Parse Error: {e}")
                    return False
            
            # Check _rels/.rels
            if '_rels/.rels' in zf.namelist():
                print(f"\n  Checking _rels/.rels...")
                with zf.open('_rels/.rels') as rels_xml:
                    rels_content = rels_xml.read()
                    try:
                        rels_tree = ET.parse(BytesIO(rels_content))
                        print(f"    ✓ Valid XML")
                    except ET.ParseError as e:
                        print(f"    ✗ XML Parse Error: {e}")
                        return False
            
        return True
        
    except zipfile.BadZipFile as e:
        print(f"  ✗ Bad ZIP file: {e}")
        return False
    except Exception as e:
        print(f"  ✗ Error: {e}")
        return False


def compare_document_xml(original: Path, exported: Path):
    """Compare document.xml between original and exported."""
    print(f"\n{'='*60}")
    print("Comparing document.xml content")
    print(f"{'='*60}")
    
    with zipfile.ZipFile(original, 'r') as zf1:
        with zf1.open('word/document.xml') as f:
            orig_content = f.read()
    
    with zipfile.ZipFile(exported, 'r') as zf2:
        with zf2.open('word/document.xml') as f:
            exp_content = f.read()
    
    print(f"\n  Original size: {len(orig_content):,} bytes")
    print(f"  Exported size: {len(exp_content):,} bytes")
    print(f"  Difference: {len(exp_content) - len(orig_content):+,} bytes")
    
    # Check encoding
    print(f"\n  Original starts with: {orig_content[:100]}")
    print(f"  Exported starts with: {exp_content[:100]}")
    
    # Check for namespace issues
    NS = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }
    
    orig_tree = ET.parse(BytesIO(orig_content))
    exp_tree = ET.parse(BytesIO(exp_content))
    
    orig_root = orig_tree.getroot()
    exp_root = exp_tree.getroot()
    
    print(f"\n  Original root tag: {orig_root.tag}")
    print(f"  Exported root tag: {exp_root.tag}")
    
    # Compare namespace declarations
    print(f"\n  Original namespaces: {len(orig_root.attrib)} attributes")
    print(f"  Exported namespaces: {len(exp_root.attrib)} attributes")
    
    # Check if namespaces are preserved
    orig_ns = set(orig_root.attrib.keys())
    exp_ns = set(exp_root.attrib.keys())
    
    missing_ns = orig_ns - exp_ns
    if missing_ns:
        print(f"\n  ✗ MISSING namespaces in export: {missing_ns}")
    else:
        print(f"\n  ✓ All namespace declarations preserved")


def test_roundtrip():
    """Test the full roundtrip: upload -> parse -> export."""
    print(f"\n{'='*60}")
    print("Testing roundtrip conversion")
    print(f"{'='*60}")
    
    from services.document_engine import docx_to_json, apply_json_to_docx
    
    # Parse original
    print("\n  Step 1: Parsing original DOCX...")
    json_doc = docx_to_json(str(UPLOAD_PATH), "test2.docx")
    
    print(f"    Blocks: {len(json_doc.blocks)}")
    print(f"    Checkboxes: {len(json_doc.checkboxes)}")
    print(f"    Dropdowns: {len(json_doc.dropdowns)}")
    
    # Show block types
    from collections import Counter
    block_types = Counter(type(b).__name__ for b in json_doc.blocks)
    print(f"    Block types: {dict(block_types)}")
    
    # Export without changes
    print("\n  Step 2: Exporting without changes...")
    test_export = Path("data/exports/debug_test.docx")
    result = apply_json_to_docx(json_doc, str(UPLOAD_PATH), str(test_export))
    print(f"    Exported to: {result}")
    
    # Check the export
    check_docx_validity(test_export, "Debug export (no changes)")
    
    return json_doc


def compare_exports_detailed(original: Path, exported: Path):
    """Compare original and exported DOCX in detail."""
    print(f"\n{'='*60}")
    print("Detailed comparison")
    print(f"{'='*60}")
    
    with zipfile.ZipFile(original, 'r') as zf1:
        with zf1.open('word/document.xml') as f:
            orig_content = f.read()
    
    with zipfile.ZipFile(exported, 'r') as zf2:
        with zf2.open('word/document.xml') as f:
            exp_content = f.read()
    
    # Check line endings
    orig_crlf = orig_content.count(b'\r\n')
    exp_crlf = exp_content.count(b'\r\n')
    orig_lf = orig_content.count(b'\n') - orig_crlf
    exp_lf = exp_content.count(b'\n') - exp_crlf
    
    print(f"\n  Line endings:")
    print(f"    Original: {orig_crlf} CRLF, {orig_lf} LF")
    print(f"    Exported: {exp_crlf} CRLF, {exp_lf} LF")
    
    # Check namespace order (first 2000 chars)
    import re
    orig_ns = re.findall(rb'xmlns:(\w+)=', orig_content[:2000])
    exp_ns = re.findall(rb'xmlns:(\w+)=', exp_content[:2000])
    
    print(f"\n  Namespace declaration order:")
    print(f"    Original: {[ns.decode() for ns in orig_ns[:10]]}...")
    print(f"    Exported: {[ns.decode() for ns in exp_ns[:10]]}...")
    
    # Check for any ns0, ns1, etc. prefixes (bad)
    bad_ns = re.findall(rb'ns\d+:', exp_content)
    if bad_ns:
        print(f"\n  ✗ BAD: Found {len(bad_ns)} occurrences of ns0/ns1/etc prefixes!")
        print(f"    Examples: {set(bad_ns[:5])}")
    else:
        print(f"\n  ✓ No bad ns0/ns1 prefixes found")


if __name__ == "__main__":
    print("DOCX Pipeline Debug Tool")
    print("=" * 60)
    
    # Check original
    orig_valid = check_docx_validity(UPLOAD_PATH, "Original upload")
    
    # Check export
    if EXPORT_PATH.exists():
        exp_valid = check_docx_validity(EXPORT_PATH, "Latest export")
        
        if orig_valid and exp_valid:
            compare_document_xml(UPLOAD_PATH, EXPORT_PATH)
    
    # Test roundtrip
    if orig_valid:
        test_roundtrip()
    
    # Detailed comparison of new export
    new_export = Path("data/exports/test2_new_export.docx")
    if new_export.exists():
        compare_exports_detailed(UPLOAD_PATH, new_export)
