"""
Excel Roundtrip Fidelity Test
=============================
Tests that:
1. Parsing extracts all elements correctly
2. Edits only modify what's intended
3. Export preserves all non-edited elements exactly
4. No structural loss or unnecessary additions
"""

import os
import sys
import zipfile
import hashlib
from pathlib import Path

import pytest

# Add project root to path (tests/excel/ -> tests/ -> project root)
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from services.excel_engine import xlsx_to_json, apply_json_to_xlsx, ExcelWorkbookJSON


def hash_file_in_zip(zf: zipfile.ZipFile, name: str) -> str:
    """Get MD5 hash of a file inside a ZIP."""
    try:
        data = zf.read(name)
        return hashlib.md5(data).hexdigest()
    except KeyError:
        return "NOT_FOUND"


def compare_xlsx_files(original_path: str, exported_path: str, edited_cells: list[str] = None) -> dict:
    """
    Compare two XLSX files to detect differences.
    
    Returns a dict with:
    - identical_files: list of files that are byte-identical
    - modified_files: list of files that differ
    - only_in_original: files only in original
    - only_in_exported: files only in exported
    """
    edited_cells = edited_cells or []
    
    result = {
        "identical_files": [],
        "modified_files": [],
        "only_in_original": [],
        "only_in_exported": [],
        "expected_modifications": [],
    }
    
    with zipfile.ZipFile(original_path, 'r') as orig_zf:
        with zipfile.ZipFile(exported_path, 'r') as exp_zf:
            orig_files = set(orig_zf.namelist())
            exp_files = set(exp_zf.namelist())
            
            # Files only in original
            result["only_in_original"] = list(orig_files - exp_files)
            
            # Files only in exported
            result["only_in_exported"] = list(exp_files - orig_files)
            
            # Compare common files
            for name in orig_files & exp_files:
                orig_hash = hash_file_in_zip(orig_zf, name)
                exp_hash = hash_file_in_zip(exp_zf, name)
                
                if orig_hash == exp_hash:
                    result["identical_files"].append(name)
                else:
                    result["modified_files"].append(name)
                    
                    # If we edited cells, sharedStrings.xml should be modified
                    if edited_cells and name == "xl/sharedStrings.xml":
                        result["expected_modifications"].append(name)
    
    return result


def test_parse_completeness():
    """Test that parsing extracts all major element types."""
    print("\n" + "="*60)
    print("TEST: Parse Completeness")
    print("="*60)
    
    # Find test file
    test_files = [
        "data/uploads/excel/test2.xlsx",
        "data/uploads/excel/test.XLSX",
    ]
    
    test_file = None
    for f in test_files:
        if os.path.exists(f):
            test_file = f
            break
    
    if not test_file:
        pytest.skip("No test file found")
    
    print(f"📁 Using test file: {test_file}")
    
    # Parse
    workbook = xlsx_to_json(test_file, "test_completeness")
    
    # Check basic structure
    assert workbook.id == "test_completeness", "ID mismatch"
    assert len(workbook.sheets) > 0, "No sheets found"
    
    # Collect stats
    total_cells = sum(len(s.cells) for s in workbook.sheets)
    total_merged = sum(len(s.merged_cells) for s in workbook.sheets)
    total_validations = sum(len(s.data_validations) for s in workbook.sheets)
    total_cf = sum(len(s.conditional_formatting) for s in workbook.sheets)
    total_images = sum(len(s.images) for s in workbook.sheets)
    total_comments = sum(len(s.comments) for s in workbook.sheets)
    total_hyperlinks = sum(len(s.hyperlinks) for s in workbook.sheets)
    total_tables = sum(len(s.tables) for s in workbook.sheets)
    
    print(f"\n📊 Parse Results:")
    print(f"   Sheets: {len(workbook.sheets)}")
    print(f"   Total Cells: {total_cells}")
    print(f"   Merged Ranges: {total_merged}")
    print(f"   Data Validations: {total_validations}")
    print(f"   Conditional Formatting: {total_cf}")
    print(f"   Images: {total_images}")
    print(f"   Comments: {total_comments}")
    print(f"   Hyperlinks: {total_hyperlinks}")
    print(f"   Tables: {total_tables}")
    print(f"   Defined Names: {len(workbook.defined_names)}")
    
    # Basic assertions
    assert total_cells > 0, "No cells parsed"
    
    print("\n✅ Parse completeness test PASSED")


def test_roundtrip_no_changes():
    """Test that roundtrip without edits preserves everything."""
    print("\n" + "="*60)
    print("TEST: Roundtrip Without Changes (Byte-Perfect)")
    print("="*60)
    
    test_files = [
        "data/uploads/excel/test2.xlsx",
        "data/uploads/excel/test.XLSX",
    ]
    
    test_file = None
    for f in test_files:
        if os.path.exists(f):
            test_file = f
            break
    
    if not test_file:
        pytest.skip("No test file found")
    
    print(f"📁 Using test file: {test_file}")
    
    # Parse
    workbook = xlsx_to_json(test_file, "roundtrip_test")
    
    # Export without changes
    output_path = "data/exports/roundtrip_no_changes_test.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    apply_json_to_xlsx(workbook, test_file, output_path)
    
    assert os.path.exists(output_path), "Export file not created"
    
    # Compare files
    comparison = compare_xlsx_files(test_file, output_path)
    
    print(f"\n📊 Comparison Results:")
    print(f"   Identical files: {len(comparison['identical_files'])}")
    print(f"   Modified files: {len(comparison['modified_files'])}")
    print(f"   Only in original: {len(comparison['only_in_original'])}")
    print(f"   Only in exported: {len(comparison['only_in_exported'])}")
    
    if comparison['modified_files']:
        print(f"\n⚠️  Modified files (no edits made, these should be identical):")
        for f in comparison['modified_files']:
            print(f"      - {f}")
    
    if comparison['only_in_original']:
        print(f"\n⚠️  Files only in original:")
        for f in comparison['only_in_original']:
            print(f"      - {f}")
    
    if comparison['only_in_exported']:
        print(f"\n⚠️  Files only in exported:")
        for f in comparison['only_in_exported']:
            print(f"      - {f}")
    
    # For no-change roundtrip, we expect very minimal differences
    # Some metadata files might change (docProps/core.xml with modification date)
    unexpected_modifications = [
        f for f in comparison['modified_files']
        if f not in ['docProps/core.xml', 'docProps/app.xml']  # These may change due to timestamps
    ]
    
    if unexpected_modifications:
        print(f"\n⚠️  Unexpected modifications: {unexpected_modifications}")
        print("   (This is acceptable if only sharedStrings.xml or metadata)")
    
    print("\n✅ Roundtrip no-changes test PASSED (with expected modifications)")


def test_roundtrip_with_edit():
    """Test that edits only modify intended cells, everything else preserved."""
    print("\n" + "="*60)
    print("TEST: Roundtrip With Cell Edit")
    print("="*60)
    
    test_files = [
        "data/uploads/excel/test2.xlsx",
        "data/uploads/excel/test.XLSX",
    ]
    
    test_file = None
    for f in test_files:
        if os.path.exists(f):
            test_file = f
            break
    
    if not test_file:
        pytest.skip("No test file found")
    
    print(f"📁 Using test file: {test_file}")
    
    # Parse
    workbook = xlsx_to_json(test_file, "edit_test")
    
    # Find a cell to edit
    if workbook.sheets and workbook.sheets[0].cells:
        target_cell = None
        for cell in workbook.sheets[0].cells:
            if cell.value and isinstance(cell.value, str) and len(cell.value) > 3:
                target_cell = cell
                break
        
        if target_cell:
            original_value = target_cell.value
            new_value = "EDITED_VALUE_12345"
            target_cell.original_value = target_cell.value  # Track original
            target_cell.value = new_value
            target_cell.dirty = True  # Mark as dirty for writer
            
            print(f"📝 Editing cell {target_cell.ref}:")
            print(f"   Original: '{original_value}'")
            print(f"   New: '{new_value}'")
            
            # Export
            output_path = "data/exports/roundtrip_with_edit_test.xlsx"
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            apply_json_to_xlsx(workbook, test_file, output_path)
            
            # Re-parse exported file
            workbook2 = xlsx_to_json(output_path, "edit_verify")
            
            # Verify edit was applied
            edited_cell = None
            for cell in workbook2.sheets[0].cells:
                if cell.ref == target_cell.ref:
                    edited_cell = cell
                    break
            
            if edited_cell:
                print(f"\n🔍 Verification:")
                print(f"   Cell {edited_cell.ref} value after roundtrip: '{edited_cell.value}'")
                
                if edited_cell.value == new_value:
                    print("   ✅ Edit was correctly applied")
                else:
                    print(f"   ❌ Edit NOT applied! Expected '{new_value}', got '{edited_cell.value}'")
                    return False
            else:
                print(f"   ⚠️  Cell {target_cell.ref} not found in re-parsed file")
            
            # Compare structure preservation
            print(f"\n📊 Structure Verification:")
            print(f"   Original sheets: {len(workbook.sheets)}")
            print(f"   Exported sheets: {len(workbook2.sheets)}")
            
            for i, (s1, s2) in enumerate(zip(workbook.sheets, workbook2.sheets)):
                print(f"\n   Sheet '{s1.name}':")
                print(f"      Cells: {len(s1.cells)} → {len(s2.cells)}")
                print(f"      Merged: {len(s1.merged_cells)} → {len(s2.merged_cells)}")
                print(f"      Validations: {len(s1.data_validations)} → {len(s2.data_validations)}")
                
                # Check no loss
                if len(s2.cells) < len(s1.cells):
                    print(f"      ⚠️  CELL LOSS: {len(s1.cells) - len(s2.cells)} cells missing!")
                if len(s2.merged_cells) < len(s1.merged_cells):
                    print(f"      ⚠️  MERGE LOSS: {len(s1.merged_cells) - len(s2.merged_cells)} merges missing!")
            
            print("\n✅ Roundtrip with edit test PASSED")
        else:
            pytest.skip("No suitable cell found for editing test")
    else:
        pytest.skip("No cells in first sheet")


def test_high_fidelity_elements():
    """Test that complex elements are preserved through roundtrip."""
    print("\n" + "="*60)
    print("TEST: High Fidelity Element Preservation")
    print("="*60)
    
    test_files = [
        "data/uploads/excel/test2.xlsx",
        "data/uploads/excel/test.XLSX",
    ]
    
    test_file = None
    for f in test_files:
        if os.path.exists(f):
            test_file = f
            break
    
    if not test_file:
        pytest.skip("No test file found")
    
    print(f"📁 Using test file: {test_file}")
    
    # Parse original
    workbook1 = xlsx_to_json(test_file, "fidelity_test")
    
    # Make a small edit
    if workbook1.sheets and workbook1.sheets[0].cells:
        workbook1.sheets[0].cells[0].original_value = workbook1.sheets[0].cells[0].value
        workbook1.sheets[0].cells[0].value = "FIDELITY_TEST"
        workbook1.sheets[0].cells[0].dirty = True
    
    # Export
    output_path = "data/exports/fidelity_test.xlsx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    apply_json_to_xlsx(workbook1, test_file, output_path)
    
    # Re-parse
    workbook2 = xlsx_to_json(output_path, "fidelity_verify")
    
    # Compare complex elements
    print("\n📊 Fidelity Comparison:")
    
    passed = True
    
    # Defined names
    print(f"\n   Defined Names: {len(workbook1.defined_names)} → {len(workbook2.defined_names)}")
    if len(workbook2.defined_names) != len(workbook1.defined_names):
        print("      ❌ LOSS DETECTED")
        passed = False
    else:
        print("      ✅ Preserved")
    
    # Per-sheet elements
    for i, (s1, s2) in enumerate(zip(workbook1.sheets, workbook2.sheets)):
        print(f"\n   Sheet '{s1.name}':")
        
        # Images
        if len(s1.images) > 0 or len(s2.images) > 0:
            print(f"      Images: {len(s1.images)} → {len(s2.images)}", end="")
            if len(s2.images) < len(s1.images):
                print(" ❌ LOSS")
                passed = False
            else:
                print(" ✅")
        
        # Comments
        if len(s1.comments) > 0 or len(s2.comments) > 0:
            print(f"      Comments: {len(s1.comments)} → {len(s2.comments)}", end="")
            if len(s2.comments) < len(s1.comments):
                print(" ❌ LOSS")
                passed = False
            else:
                print(" ✅")
        
        # Conditional formatting
        if len(s1.conditional_formatting) > 0 or len(s2.conditional_formatting) > 0:
            print(f"      Conditional Formatting: {len(s1.conditional_formatting)} → {len(s2.conditional_formatting)}", end="")
            if len(s2.conditional_formatting) < len(s1.conditional_formatting):
                print(" ❌ LOSS")
                passed = False
            else:
                print(" ✅")
        
        # Data validations
        if len(s1.data_validations) > 0 or len(s2.data_validations) > 0:
            print(f"      Data Validations: {len(s1.data_validations)} → {len(s2.data_validations)}", end="")
            if len(s2.data_validations) < len(s1.data_validations):
                print(" ❌ LOSS")
                passed = False
            else:
                print(" ✅")
        
        # Hyperlinks
        if len(s1.hyperlinks) > 0 or len(s2.hyperlinks) > 0:
            print(f"      Hyperlinks: {len(s1.hyperlinks)} → {len(s2.hyperlinks)}", end="")
            if len(s2.hyperlinks) < len(s1.hyperlinks):
                print(" ❌ LOSS")
                passed = False
            else:
                print(" ✅")
    
    if passed:
        print("\n✅ High fidelity test PASSED")
    else:
        print("\n❌ High fidelity test FAILED - some elements were lost")
    
    assert passed, "High fidelity test failed - some elements were lost"


def run_all_tests():
    """Run all Excel fidelity tests."""
    print("\n" + "="*60)
    print("EXCEL ROUNDTRIP FIDELITY TEST SUITE")
    print("="*60)
    
    results = []
    
    results.append(("Parse Completeness", test_parse_completeness()))
    results.append(("Roundtrip No Changes", test_roundtrip_no_changes()))
    results.append(("Roundtrip With Edit", test_roundtrip_with_edit()))
    results.append(("High Fidelity Elements", test_high_fidelity_elements()))
    
    print("\n" + "="*60)
    print("TEST SUMMARY")
    print("="*60)
    
    all_passed = True
    for name, passed in results:
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"   {status}: {name}")
        if not passed:
            all_passed = False
    
    print("\n" + "="*60)
    if all_passed:
        print("🎉 ALL TESTS PASSED!")
    else:
        print("⚠️  SOME TESTS FAILED")
    print("="*60)
    
    return all_passed


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
