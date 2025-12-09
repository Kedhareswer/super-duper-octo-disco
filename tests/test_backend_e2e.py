"""End-to-end backend test for the document pipeline.

Tests:
1. Upload DOCX → Parse to JSON
2. Verify JSON structure integrity
3. Modify JSON (simulate edits)
4. Export back to DOCX
5. Verify exported DOCX validity
6. Re-parse exported DOCX and compare
"""

import json
import os
import shutil
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

# Add project root to path
import sys
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from services.document_engine import docx_to_json, apply_json_to_docx, validate_document_json
from models.schemas import DocumentJSON, ParagraphBlock, TableBlock, DrawingBlock


DEFAULT_TEST_FILE = PROJECT_ROOT / "data/uploads/docx/test2.docx"
TEST_OUTPUT_DIR = PROJECT_ROOT / "data/test_outputs"

# Parse command line for custom test file
TEST_FILE = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_TEST_FILE


def print_header(title: str):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")


def print_result(success: bool, message: str):
    icon = "✓" if success else "✗"
    print(f"  {icon} {message}")


def test_1_upload_and_parse():
    """Test 1: Upload and parse DOCX to JSON."""
    print_header("TEST 1: Upload and Parse DOCX")
    
    if not TEST_FILE.exists():
        print_result(False, f"Test file not found: {TEST_FILE}")
        return None
    
    print_result(True, f"Test file exists: {TEST_FILE} ({TEST_FILE.stat().st_size:,} bytes)")
    
    try:
        json_doc = docx_to_json(str(TEST_FILE), "test2.docx")
        print_result(True, f"Parsed successfully")
        print(f"\n  Document structure:")
        print(f"    - ID: {json_doc.id}")
        print(f"    - Title: {json_doc.title}")
        print(f"    - Blocks: {len(json_doc.blocks)}")
        print(f"    - Checkboxes: {len(json_doc.checkboxes)}")
        print(f"    - Dropdowns: {len(json_doc.dropdowns)}")
        
        # Count block types
        para_count = sum(1 for b in json_doc.blocks if isinstance(b, ParagraphBlock))
        table_count = sum(1 for b in json_doc.blocks if isinstance(b, TableBlock))
        drawing_count = sum(1 for b in json_doc.blocks if isinstance(b, DrawingBlock))
        
        print(f"\n  Block breakdown:")
        print(f"    - Paragraphs: {para_count}")
        print(f"    - Tables: {table_count}")
        print(f"    - Drawings: {drawing_count}")
        
        return json_doc
        
    except Exception as e:
        print_result(False, f"Parse failed: {e}")
        import traceback
        traceback.print_exc()
        return None


def test_2_json_structure(json_doc: DocumentJSON):
    """Test 2: Verify JSON structure integrity."""
    print_header("TEST 2: Verify JSON Structure")
    
    if json_doc is None:
        print_result(False, "No document to test")
        return False
    
    # Validate using built-in validator
    validation = validate_document_json(json_doc)
    if validation.is_valid:
        print_result(True, "Document passes validation")
    else:
        print_result(False, f"Validation failed with {len(validation.errors)} errors:")
        for err in validation.errors[:5]:
            print(f"      - {err.field}: {err.message}")
        return False
    
    # Check all blocks have xml_ref
    missing_refs = []
    for i, block in enumerate(json_doc.blocks):
        if not block.xml_ref:
            missing_refs.append(f"block[{i}]")
        
        if isinstance(block, ParagraphBlock):
            for j, run in enumerate(block.runs):
                if not run.xml_ref:
                    missing_refs.append(f"block[{i}].runs[{j}]")
        
        elif isinstance(block, TableBlock):
            for ri, row in enumerate(block.rows):
                for ci, cell in enumerate(row.cells):
                    if not cell.xml_ref:
                        missing_refs.append(f"block[{i}].rows[{ri}].cells[{ci}]")
    
    if missing_refs:
        print_result(False, f"Missing xml_ref in {len(missing_refs)} elements")
        for ref in missing_refs[:5]:
            print(f"      - {ref}")
    else:
        print_result(True, "All elements have xml_ref")
    
    # Check text content extraction
    total_text = 0
    for block in json_doc.blocks:
        if isinstance(block, ParagraphBlock):
            for run in block.runs:
                if run.text:
                    total_text += len(run.text)
        elif isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.blocks:
                        for run in para.runs:
                            if run.text:
                                total_text += len(run.text)
    
    print_result(True, f"Total text extracted: {total_text} characters")
    
    # Check serialization round-trip
    try:
        json_str = json_doc.model_dump_json()
        restored = DocumentJSON.model_validate_json(json_str)
        print_result(True, f"JSON serialization round-trip OK ({len(json_str):,} bytes)")
    except Exception as e:
        print_result(False, f"JSON serialization failed: {e}")
        return False
    
    return True


def test_3_modify_json(json_doc: DocumentJSON):
    """Test 3: Modify JSON (simulate edits)."""
    print_header("TEST 3: Modify JSON (Simulate Edits)")
    
    if json_doc is None:
        print_result(False, "No document to test")
        return None
    
    # Deep copy to avoid modifying original
    modified = DocumentJSON.model_validate_json(json_doc.model_dump_json())
    
    modifications = []
    
    # Find and modify first paragraph with text
    for block in modified.blocks:
        if isinstance(block, ParagraphBlock) and block.runs:
            for run in block.runs:
                if run.text and len(run.text) > 5:
                    original = run.text
                    run.text = run.text + " [EDITED]"
                    modifications.append(f"Paragraph: '{original[:30]}...' → added '[EDITED]'")
                    break
            if modifications:
                break
    
    # Find and modify first table cell
    for block in modified.blocks:
        if isinstance(block, TableBlock):
            for row in block.rows:
                for cell in row.cells:
                    if cell.blocks:
                        for para in cell.blocks:
                            for run in para.runs:
                                if run.text and len(run.text) > 3:
                                    original = run.text
                                    run.text = run.text.upper()
                                    modifications.append(f"Table cell: '{original[:20]}' → '{run.text[:20]}'")
                                    break
                            if len(modifications) >= 2:
                                break
                        if len(modifications) >= 2:
                            break
                    if len(modifications) >= 2:
                        break
                if len(modifications) >= 2:
                    break
            if len(modifications) >= 2:
                break
    
    # Toggle a checkbox if present
    if modified.checkboxes:
        cb = modified.checkboxes[0]
        original_state = cb.checked
        cb.checked = not cb.checked
        modifications.append(f"Checkbox '{cb.label}': {original_state} → {cb.checked}")
    
    # Change dropdown selection if present
    if modified.dropdowns:
        dd = modified.dropdowns[0]
        if dd.options and len(dd.options) > 1:
            original = dd.selected
            # Pick a different option
            for opt in dd.options:
                if opt != dd.selected:
                    dd.selected = opt
                    modifications.append(f"Dropdown '{dd.label}': '{original}' → '{dd.selected}'")
                    break
    
    if modifications:
        print_result(True, f"Made {len(modifications)} modifications:")
        for mod in modifications:
            print(f"      - {mod}")
    else:
        print_result(False, "Could not find any content to modify")
    
    return modified


def test_4_export_docx(json_doc: DocumentJSON, label: str = "modified"):
    """Test 4: Export JSON back to DOCX."""
    print_header(f"TEST 4: Export to DOCX ({label})")
    
    if json_doc is None:
        print_result(False, "No document to export")
        return None
    
    TEST_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = TEST_OUTPUT_DIR / f"test2_{label}.docx"
    
    try:
        result_path = apply_json_to_docx(
            json_doc=json_doc,
            base_docx_path=str(TEST_FILE),
            out_docx_path=str(output_path)
        )
        print_result(True, f"Exported to: {result_path}")
        print_result(True, f"File size: {Path(result_path).stat().st_size:,} bytes")
        return Path(result_path)
        
    except Exception as e:
        print_result(False, f"Export failed: {e}")
        import traceback
        traceback.print_exc()
        return None


def test_5_verify_docx(docx_path: Path):
    """Test 5: Verify exported DOCX validity."""
    print_header("TEST 5: Verify Exported DOCX")
    
    if docx_path is None or not docx_path.exists():
        print_result(False, "No file to verify")
        return False
    
    issues = []
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            # Check required files
            required = ['word/document.xml', '[Content_Types].xml', '_rels/.rels']
            for req in required:
                if req in zf.namelist():
                    print_result(True, f"Required file present: {req}")
                else:
                    print_result(False, f"Missing required file: {req}")
                    issues.append(f"Missing {req}")
            
            # Parse and validate document.xml
            with zf.open('word/document.xml') as f:
                content = f.read()
            
            # Check XML declaration
            if b'standalone="yes"' in content:
                print_result(True, "XML declaration has standalone='yes'")
            else:
                print_result(False, "Missing standalone='yes' in XML declaration")
                issues.append("Missing standalone='yes'")
            
            # Check for bad namespace prefixes
            import re
            bad_ns = re.findall(rb'<ns\d+:', content)
            if bad_ns:
                print_result(False, f"Found {len(bad_ns)} bad namespace prefixes (ns0:, ns1:, etc.)")
                issues.append(f"Bad namespace prefixes: {len(bad_ns)}")
            else:
                print_result(True, "No bad namespace prefixes")
            
            # Check for w: prefix
            if b'<w:' in content:
                print_result(True, "Uses correct w: namespace prefix")
            else:
                print_result(False, "Missing w: namespace prefix")
                issues.append("Missing w: prefix")
            
            # Try to parse as XML
            try:
                tree = ET.parse(BytesIO(content))
                root = tree.getroot()
                print_result(True, f"Valid XML structure")
                
                # Count elements
                NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                body = root.find("w:body", NS)
                if body is not None:
                    paras = body.findall(".//w:p", NS)
                    tables = body.findall(".//w:tbl", NS)
                    print_result(True, f"Contains {len(paras)} paragraphs, {len(tables)} tables")
                else:
                    print_result(False, "Missing w:body element")
                    issues.append("Missing w:body")
                    
            except ET.ParseError as e:
                print_result(False, f"XML parse error: {e}")
                issues.append(f"XML parse error: {e}")
            
    except zipfile.BadZipFile as e:
        print_result(False, f"Invalid ZIP file: {e}")
        issues.append(f"Bad ZIP: {e}")
    except Exception as e:
        print_result(False, f"Error: {e}")
        issues.append(str(e))
    
    return len(issues) == 0


def test_6_reparse_and_compare(original_json: DocumentJSON, exported_path: Path):
    """Test 6: Re-parse exported DOCX and compare."""
    print_header("TEST 6: Re-parse and Compare")
    
    if original_json is None or exported_path is None:
        print_result(False, "Missing inputs for comparison")
        return False
    
    try:
        reparsed = docx_to_json(str(exported_path), "reparsed")
        print_result(True, "Re-parsed exported DOCX successfully")
        
        # Compare structure
        orig_blocks = len(original_json.blocks)
        new_blocks = len(reparsed.blocks)
        
        if orig_blocks == new_blocks:
            print_result(True, f"Block count preserved: {orig_blocks}")
        else:
            print_result(False, f"Block count changed: {orig_blocks} → {new_blocks}")
        
        # Compare block types
        def count_types(doc):
            return {
                'para': sum(1 for b in doc.blocks if isinstance(b, ParagraphBlock)),
                'table': sum(1 for b in doc.blocks if isinstance(b, TableBlock)),
                'drawing': sum(1 for b in doc.blocks if isinstance(b, DrawingBlock)),
            }
        
        orig_types = count_types(original_json)
        new_types = count_types(reparsed)
        
        if orig_types == new_types:
            print_result(True, f"Block types preserved: {orig_types}")
        else:
            print_result(False, f"Block types changed: {orig_types} → {new_types}")
        
        # Compare checkboxes
        if len(original_json.checkboxes) == len(reparsed.checkboxes):
            print_result(True, f"Checkbox count preserved: {len(original_json.checkboxes)}")
        else:
            print_result(False, f"Checkbox count changed: {len(original_json.checkboxes)} → {len(reparsed.checkboxes)}")
        
        # Compare dropdowns
        if len(original_json.dropdowns) == len(reparsed.dropdowns):
            print_result(True, f"Dropdown count preserved: {len(original_json.dropdowns)}")
        else:
            print_result(False, f"Dropdown count changed: {len(original_json.dropdowns)} → {len(reparsed.dropdowns)}")
        
        # Compare text content
        def extract_all_text(doc):
            texts = []
            for block in doc.blocks:
                if isinstance(block, ParagraphBlock):
                    for run in block.runs:
                        if run.text:
                            texts.append(run.text)
                elif isinstance(block, TableBlock):
                    for row in block.rows:
                        for cell in row.cells:
                            for para in cell.blocks:
                                for run in para.runs:
                                    if run.text:
                                        texts.append(run.text)
            return texts
        
        orig_texts = extract_all_text(original_json)
        new_texts = extract_all_text(reparsed)
        
        # Note: text might be modified, so just check count
        print(f"\n  Text comparison:")
        print(f"    Original text segments: {len(orig_texts)}")
        print(f"    Re-parsed text segments: {len(new_texts)}")
        
        # Show sample differences
        if orig_texts and new_texts:
            print(f"\n  Sample text (first 3 segments):")
            for i in range(min(3, len(orig_texts))):
                orig = orig_texts[i][:50] if i < len(orig_texts) else "(missing)"
                new = new_texts[i][:50] if i < len(new_texts) else "(missing)"
                match = "✓" if orig == new else "≠"
                print(f"    {match} '{orig}' vs '{new}'")
        
        return True
        
    except Exception as e:
        print_result(False, f"Re-parse failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_7_compare_with_original_docx(exported_path: Path):
    """Test 7: Compare exported DOCX with original at byte level."""
    print_header("TEST 7: Compare with Original DOCX")
    
    if exported_path is None or not exported_path.exists():
        print_result(False, "No exported file to compare")
        return
    
    with zipfile.ZipFile(TEST_FILE, 'r') as orig_zf:
        with zipfile.ZipFile(exported_path, 'r') as exp_zf:
            orig_files = set(orig_zf.namelist())
            exp_files = set(exp_zf.namelist())
            
            # Check file list
            if orig_files == exp_files:
                print_result(True, f"Same files in archive: {len(orig_files)}")
            else:
                missing = orig_files - exp_files
                extra = exp_files - orig_files
                if missing:
                    print_result(False, f"Missing files: {missing}")
                if extra:
                    print_result(False, f"Extra files: {extra}")
            
            # Compare file sizes
            print(f"\n  File size comparison:")
            for name in sorted(orig_files & exp_files):
                orig_size = orig_zf.getinfo(name).file_size
                exp_size = exp_zf.getinfo(name).file_size
                diff = exp_size - orig_size
                if diff == 0:
                    status = "="
                elif abs(diff) < 100:
                    status = f"+{diff}" if diff > 0 else str(diff)
                else:
                    status = f"+{diff:,}" if diff > 0 else f"{diff:,}"
                
                if name == "word/document.xml":
                    print(f"    {name}: {orig_size:,} → {exp_size:,} ({status}) ← MAIN")
                elif diff != 0:
                    print(f"    {name}: {orig_size:,} → {exp_size:,} ({status})")


def run_all_tests():
    """Run all tests in sequence."""
    print("\n" + "="*70)
    print("  BACKEND END-TO-END TEST SUITE")
    print("  Testing: " + str(TEST_FILE))
    print("="*70)
    
    # Test 1: Upload and parse
    json_doc = test_1_upload_and_parse()
    if json_doc is None:
        print("\n❌ FAILED: Cannot continue without parsed document")
        return
    
    # Test 2: Verify JSON structure
    if not test_2_json_structure(json_doc):
        print("\n⚠️ WARNING: JSON structure has issues")
    
    # Test 3: Modify JSON
    modified_doc = test_3_modify_json(json_doc)
    
    # Test 4a: Export unmodified (baseline)
    unmodified_path = test_4_export_docx(json_doc, "unmodified")
    
    # Test 4b: Export modified
    modified_path = test_4_export_docx(modified_doc, "modified") if modified_doc else None
    
    # Test 5: Verify exported DOCX
    if unmodified_path:
        test_5_verify_docx(unmodified_path)
    
    # Test 6: Re-parse and compare
    if unmodified_path:
        test_6_reparse_and_compare(json_doc, unmodified_path)
    
    # Test 7: Compare with original
    if unmodified_path:
        test_7_compare_with_original_docx(unmodified_path)
    
    print("\n" + "="*70)
    print("  TEST SUITE COMPLETE")
    print("="*70)
    print(f"\n  Output files in: {TEST_OUTPUT_DIR.absolute()}")
    if unmodified_path:
        print(f"  - {unmodified_path.name} (no changes, for Word validation)")
    if modified_path:
        print(f"  - {modified_path.name} (with edits)")
    print("\n  → Open these files in Microsoft Word to verify they work correctly.")


if __name__ == "__main__":
    run_all_tests()
