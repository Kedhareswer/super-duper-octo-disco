"""Debug Excel cell values - inspect parsed cell data."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from services.excel_engine import xlsx_to_json

TEST_FILE = Path("data/uploads/excel/excel_test.XLSX")


def main():
    if not TEST_FILE.exists():
        print(f"Test file not found: {TEST_FILE}")
        return
    
    print(f"Parsing: {TEST_FILE}")
    print("=" * 60)
    
    wb = xlsx_to_json(str(TEST_FILE), "debug-cells")
    
    print(f"\nWorkbook: {wb.id}")
    print(f"Sheets: {len(wb.sheets)}")
    print(f"Shared strings: {len(wb.shared_strings)}")
    
    for i, sheet in enumerate(wb.sheets):
        print(f"\n{'='*60}")
        print(f"Sheet: {sheet.name} (index {i})")
        print(f"  Cells: {len(sheet.cells)}")
        print(f"  Merged ranges: {len(sheet.merged_cells)}")
        print(f"  Data validations: {len(sheet.data_validations)}")
        print(f"  Images: {len(sheet.images)}")
        print(f"  Comments: {len(sheet.comments)}")
        
        # Show first 10 cells
        print(f"\n  First 10 cells:")
        for cell in sheet.cells[:10]:
            val_preview = str(cell.value)[:40] if cell.value else "(empty)"
            print(f"    {cell.ref}: {cell.data_type.value if cell.data_type else 'unknown'} = {val_preview}")
        
        # Show merged cells
        if sheet.merged_cells:
            print(f"\n  Merged ranges:")
            for mc in sheet.merged_cells[:5]:
                print(f"    {mc.ref}")
        
        # Show data validations (dropdowns)
        if sheet.data_validations:
            print(f"\n  Data validations:")
            for dv in sheet.data_validations[:5]:
                print(f"    {dv.sqref}: type={dv.type}, formula={dv.formula1[:30] if dv.formula1 else 'N/A'}...")


if __name__ == "__main__":
    main()
