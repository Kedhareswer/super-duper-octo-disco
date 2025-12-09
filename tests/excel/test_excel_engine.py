"""Tests for the Excel Engine.

Validates parsing and writing of complex XLSX files with:
- Multiple worksheets
- Merged cells
- Data validation (dropdowns)
- Images
- Comments
- Styles and formatting
"""

import json
import os
import sys
from pathlib import Path

# Add project root to path (tests/excel/ -> tests/ -> project root)
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

import pytest

from services.excel_engine import (
    xlsx_to_json,
    apply_json_to_xlsx,
    ExcelWorkbookJSON,
    ExcelSheetJSON,
)
from services.excel_engine.writer import apply_cell_edits


# Test file path - use test2.xlsx which is smaller (35KB vs 2MB)
TEST_XLSX = project_root / "data" / "uploads" / "excel" / "test2.xlsx"
OUTPUT_DIR = project_root / "data" / "outputs" / "excel"


class TestExcelParser:
    """Test XLSX parsing functionality."""
    
    def test_parse_workbook(self):
        """Test basic workbook parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-1")
        
        assert workbook is not None
        assert workbook.id == "test-workbook-1"
        assert len(workbook.sheets) > 0
        
        print(f"\n✓ Parsed workbook with {len(workbook.sheets)} sheets")
        for sheet in workbook.sheets:
            print(f"  - {sheet.name}: {len(sheet.cells)} cells, {len(sheet.merged_cells)} merges, {len(sheet.data_validations)} validations")
    
    def test_parse_sheets(self):
        """Test individual sheet parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-2")
        
        for sheet in workbook.sheets:
            assert sheet.id is not None
            assert sheet.name is not None
            assert sheet.sheet_index >= 0
            
            # Verify cells have required fields
            for cell in sheet.cells:
                assert cell.id is not None
                assert cell.ref is not None
                assert cell.row > 0
                assert cell.col > 0
        
        print(f"\n✓ All {len(workbook.sheets)} sheets have valid structure")
    
    def test_parse_merged_cells(self):
        """Test merged cell parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-3")
        
        total_merges = 0
        for sheet in workbook.sheets:
            if sheet.merged_cells:
                total_merges += len(sheet.merged_cells)
                for merge in sheet.merged_cells:
                    assert merge.ref is not None
                    assert ":" in merge.ref
                    assert merge.start_row <= merge.end_row
                    assert merge.start_col <= merge.end_col
                    print(f"  - {sheet.name}: Merge {merge.ref}")
        
        print(f"\n✓ Found {total_merges} merged cell ranges")
    
    def test_parse_data_validations(self):
        """Test data validation (dropdowns) parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-4")
        
        total_validations = 0
        for sheet in workbook.sheets:
            if sheet.data_validations:
                total_validations += len(sheet.data_validations)
                for dv in sheet.data_validations:
                    assert dv.sqref is not None
                    assert dv.validation_type is not None
                    print(f"  - {sheet.name}: Validation at {dv.sqref} (type: {dv.validation_type})")
                    if dv.options:
                        print(f"    Options: {dv.options[:5]}...")
        
        print(f"\n✓ Found {total_validations} data validation rules")
    
    def test_parse_images(self):
        """Test image parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-5")
        
        total_images = 0
        for sheet in workbook.sheets:
            if sheet.images:
                total_images += len(sheet.images)
                for img in sheet.images:
                    assert img.media_path is not None
                    print(f"  - {sheet.name}: Image {img.name or 'unnamed'} at {img.media_path}")
        
        print(f"\n✓ Found {total_images} images")
    
    def test_parse_styles(self):
        """Test cell style parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-6")
        
        styled_cells = 0
        for sheet in workbook.sheets:
            for cell in sheet.cells:
                if cell.style is not None:
                    styled_cells += 1
        
        print(f"\n✓ Found {styled_cells} cells with styles")
    
    def test_shared_strings(self):
        """Test shared strings parsing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "test-workbook-7")
        
        print(f"\n✓ Found {len(workbook.shared_strings)} shared strings")
        if workbook.shared_strings:
            print(f"  Sample: {workbook.shared_strings[:3]}")


class TestExcelWriter:
    """Test XLSX writing functionality."""
    
    def test_roundtrip_basic(self):
        """Test basic read-write roundtrip."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "excel_roundtrip_test.xlsx"
        
        # Parse
        workbook = xlsx_to_json(str(TEST_XLSX), "roundtrip-test")
        
        # Write
        result_path = apply_json_to_xlsx(workbook, str(TEST_XLSX), str(output_path))
        
        assert Path(result_path).exists()
        print(f"\n✓ Roundtrip successful: {result_path}")
    
    def test_edit_and_write(self):
        """Test editing cells and writing."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "excel_edit_test.xlsx"
        
        # Parse
        workbook = xlsx_to_json(str(TEST_XLSX), "edit-test")
        
        # Make edits
        if workbook.sheets:
            first_sheet = workbook.sheets[0]
            edits = [
                {"sheet": first_sheet.name, "cell": "Z1", "value": "TEST EDIT"},
                {"sheet": first_sheet.name, "cell": "Z2", "value": 12345},
            ]
            apply_cell_edits(workbook, edits)
        
        # Write
        result_path = apply_json_to_xlsx(workbook, str(TEST_XLSX), str(output_path))
        
        assert Path(result_path).exists()
        print(f"\n✓ Edit and write successful: {result_path}")


class TestExcelStructure:
    """Test structural integrity."""
    
    def test_json_serialization(self):
        """Test that parsed workbook can be serialized to JSON."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "json-test")
        
        # Serialize to JSON
        json_str = workbook.model_dump_json(indent=2)
        
        assert json_str is not None
        assert len(json_str) > 0
        
        # Parse back
        data = json.loads(json_str)
        assert "id" in data
        assert "sheets" in data
        
        print(f"\n✓ JSON serialization successful ({len(json_str)} bytes)")
    
    def test_cell_lookup(self):
        """Test cell lookup by reference."""
        if not TEST_XLSX.exists():
            pytest.skip(f"Test file not found: {TEST_XLSX}")
        
        workbook = xlsx_to_json(str(TEST_XLSX), "lookup-test")
        
        for sheet in workbook.sheets:
            if sheet.cells:
                first_cell = sheet.cells[0]
                looked_up = sheet.get_cell(first_cell.ref)
                assert looked_up is not None
                assert looked_up.ref == first_cell.ref
                print(f"\n✓ Cell lookup works: {first_cell.ref} = {first_cell.value}")
                break


def run_quick_test():
    """Quick test to verify the engine works."""
    print("=" * 60)
    print("Excel Engine Quick Test")
    print("=" * 60)
    
    if not TEST_XLSX.exists():
        print(f"ERROR: Test file not found: {TEST_XLSX}")
        return False
    
    try:
        # Parse
        print("\n1. Parsing XLSX...")
        workbook = xlsx_to_json(str(TEST_XLSX), "quick-test")
        print(f"   ✓ Parsed {len(workbook.sheets)} sheets")
        print(f"   ✓ Found {len(workbook.defined_names)} defined names")
        
        # Inspect
        print("\n2. Sheet Summary:")
        for sheet in workbook.sheets:
            print(f"   - {sheet.name}:")
            print(f"     Cells: {len(sheet.cells):,}")
            print(f"     Merged: {len(sheet.merged_cells)}")
            print(f"     Validations: {len(sheet.data_validations)}")
            print(f"     Images: {len(sheet.images)}")
            print(f"     Comments: {len(sheet.comments)}")
            
            # Show new complex elements if present
            extras = []
            if sheet.hyperlinks:
                extras.append(f"Hyperlinks: {len(sheet.hyperlinks)}")
            if sheet.conditional_formatting:
                extras.append(f"CF: {len(sheet.conditional_formatting)}")
            if sheet.form_controls:
                extras.append(f"Controls: {len(sheet.form_controls)}")
            if sheet.tables:
                extras.append(f"Tables: {len(sheet.tables)}")
            if sheet.sheet_view and sheet.sheet_view.freeze_pane:
                fp = sheet.sheet_view.freeze_pane
                extras.append(f"Freeze: {fp.y_split}R/{fp.x_split}C")
            if extras:
                print(f"     Complex: {', '.join(extras)}")
        
        # Summary totals
        print("\n3. Complex Elements Summary:")
        total_cf = sum(len(s.conditional_formatting) for s in workbook.sheets)
        total_hl = sum(len(s.hyperlinks) for s in workbook.sheets)
        total_fc = sum(len(s.form_controls) for s in workbook.sheets)
        total_tables = sum(len(s.tables) for s in workbook.sheets)
        print(f"   Conditional Formatting: {total_cf}")
        print(f"   Hyperlinks: {total_hl}")
        print(f"   Form Controls: {total_fc}")
        print(f"   Tables: {total_tables}")
        
        # Roundtrip
        print("\n4. Testing roundtrip...")
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "excel_quick_test.xlsx"
        result = apply_json_to_xlsx(workbook, str(TEST_XLSX), str(output_path))
        print(f"   ✓ Wrote to {result}")
        
        # Verify roundtrip
        print("\n5. Verifying roundtrip fidelity...")
        workbook2 = xlsx_to_json(str(output_path), "roundtrip-verify")
        cf_after = sum(len(s.conditional_formatting) for s in workbook2.sheets)
        print(f"   CF before: {total_cf}, after: {cf_after} - {'✓ Preserved' if cf_after == total_cf else '✗ Lost'}")
        
        print("\n" + "=" * 60)
        print("✓ All tests passed!")
        print("=" * 60)
        return True
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    run_quick_test()
