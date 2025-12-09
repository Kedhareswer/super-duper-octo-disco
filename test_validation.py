"""Comprehensive validation test for all three test documents."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from services.document_engine import docx_to_json, apply_json_to_docx
from services.validation import (
    validate_parse_stage,
    validate_full_roundtrip,
    extract_raw_docx_content,
    extract_json_content,
    print_report,
)

TEST_FILES = [
    "data/uploads/docx/test.docx",
    "data/uploads/docx/test2.docx",
    "data/uploads/docx/test3.DOCX",
]

def test_parse_stage(docx_path: str):
    """Test Stage 1: DOCX → JSON parsing."""
    print(f"\n{'#'*70}")
    print(f"# STAGE 1: PARSE TEST - {Path(docx_path).name}")
    print(f"{'#'*70}")
    
    # Parse
    json_doc = docx_to_json(docx_path, Path(docx_path).name)
    
    # Validate
    report = validate_parse_stage(docx_path, json_doc)
    print_report(report)
    
    return json_doc, report


def test_roundtrip(docx_path: str):
    """Test Full Roundtrip: DOCX → JSON → DOCX (no edits)."""
    print(f"\n{'#'*70}")
    print(f"# ROUNDTRIP TEST - {Path(docx_path).name}")
    print(f"{'#'*70}")
    
    # Parse
    json_doc = docx_to_json(docx_path, Path(docx_path).name)
    
    # Export without changes
    output_path = f"data/test_outputs/roundtrip_{Path(docx_path).name}"
    Path("data/test_outputs").mkdir(parents=True, exist_ok=True)
    
    apply_json_to_docx(json_doc, docx_path, output_path)
    
    # Validate roundtrip
    report = validate_full_roundtrip(docx_path, json_doc, output_path)
    print_report(report)
    
    return report


def test_xml_ref_resolution(docx_path: str):
    """Test that xml_ref paths correctly resolve to elements."""
    print(f"\n{'#'*70}")
    print(f"# XML_REF RESOLUTION TEST - {Path(docx_path).name}")
    print(f"{'#'*70}")
    
    import zipfile
    from xml.etree import ElementTree as ET
    from services.document_engine import docx_to_json, _find_node_by_ref, NS
    
    # Parse
    json_doc = docx_to_json(docx_path, Path(docx_path).name)
    
    # Load DOCX XML
    with zipfile.ZipFile(docx_path, 'r') as zf:
        with zf.open('word/document.xml') as doc:
            tree = ET.parse(doc)
            root = tree.getroot()
            body = root.find("w:body", NS)
    
    # Test resolution
    total_refs = 0
    resolved = 0
    failed = []
    
    for block in json_doc.blocks:
        if block.type.value == "paragraph":
            total_refs += 1
            el = _find_node_by_ref(body, block.xml_ref)
            if el is not None:
                resolved += 1
            else:
                failed.append(block.xml_ref)
        
        elif block.type.value == "table":
            total_refs += 1
            el = _find_node_by_ref(body, block.xml_ref)
            if el is not None:
                resolved += 1
            else:
                failed.append(block.xml_ref)
            
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.blocks:
                        total_refs += 1
                        el = _find_node_by_ref(body, para.xml_ref)
                        if el is not None:
                            resolved += 1
                        else:
                            failed.append(para.xml_ref)
    
    print(f"\nXML Reference Resolution:")
    print(f"  Total refs: {total_refs}")
    print(f"  Resolved: {resolved}")
    print(f"  Failed: {len(failed)}")
    
    if failed:
        print(f"\n  Failed refs (first 10):")
        for ref in failed[:10]:
            print(f"    - {ref}")
    else:
        print(f"\n  ✓ All references resolved correctly")
    
    return len(failed) == 0


def test_text_patching(docx_path: str):
    """Test that text patching works correctly."""
    print(f"\n{'#'*70}")
    print(f"# TEXT PATCHING TEST - {Path(docx_path).name}")
    print(f"{'#'*70}")
    
    # Parse
    json_doc = docx_to_json(docx_path, Path(docx_path).name)
    
    # Make a simple edit - find a text run and modify it
    test_edit_made = False
    original_text = None
    new_text = "[EDITED BY TEST]"
    edited_block_ref = None
    
    for block in json_doc.blocks:
        if block.type.value == "paragraph" and block.runs:
            for run in block.runs:
                if run.text and len(run.text) > 5 and "click" not in run.text.lower():
                    original_text = run.text
                    run.text = new_text
                    edited_block_ref = block.xml_ref
                    test_edit_made = True
                    break
            if test_edit_made:
                break
    
    if not test_edit_made:
        print("  Could not find suitable text to edit")
        return False
    
    print(f"  Original text: '{original_text[:50]}...'")
    print(f"  New text: '{new_text}'")
    print(f"  Block ref: {edited_block_ref}")
    
    # Export
    output_path = f"data/test_outputs/patched_{Path(docx_path).name}"
    apply_json_to_docx(json_doc, docx_path, output_path)
    
    # Re-parse output to verify
    output_json = docx_to_json(output_path, Path(output_path).name)
    
    # Find the edited text
    found_edit = False
    for block in output_json.blocks:
        if block.type.value == "paragraph":
            for run in block.runs:
                if run.text == new_text:
                    found_edit = True
                    break
        if found_edit:
            break
    
    if found_edit:
        print(f"\n  ✓ Edit was successfully applied and preserved")
    else:
        print(f"\n  ✗ Edit was NOT found in output!")
    
    return found_edit


def main():
    print("=" * 70)
    print("COMPREHENSIVE DOCUMENT PROCESSING VALIDATION")
    print("=" * 70)
    
    all_passed = True
    
    for docx_path in TEST_FILES:
        if not Path(docx_path).exists():
            print(f"\n⚠ File not found: {docx_path}")
            continue
        
        print(f"\n\n{'*'*70}")
        print(f"* TESTING: {docx_path}")
        print(f"{'*'*70}")
        
        # Test 1: Parse stage
        json_doc, parse_report = test_parse_stage(docx_path)
        if parse_report.has_errors:
            all_passed = False
        
        # Test 2: XML ref resolution
        if not test_xml_ref_resolution(docx_path):
            all_passed = False
        
        # Test 3: Roundtrip
        roundtrip_report = test_roundtrip(docx_path)
        if roundtrip_report.has_errors:
            all_passed = False
        
        # Test 4: Text patching
        if not test_text_patching(docx_path):
            all_passed = False
    
    print(f"\n\n{'='*70}")
    print(f"FINAL RESULT: {'ALL TESTS PASSED' if all_passed else 'SOME TESTS FAILED'}")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()
