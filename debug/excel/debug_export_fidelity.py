"""Debug Excel export fidelity - compare original vs exported XLSX."""

import sys
import zipfile
import hashlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from services.excel_engine import xlsx_to_json, apply_json_to_xlsx

TEST_FILE = Path("data/uploads/excel/excel_test.XLSX")
OUTPUT_FILE = Path("data/outputs/excel/debug_export.xlsx")


def hash_content(data: bytes) -> str:
    """Get MD5 hash of content."""
    return hashlib.md5(data).hexdigest()[:12]


def compare_xlsx(original: Path, exported: Path):
    """Compare two XLSX files and report differences."""
    
    print(f"\nComparing:")
    print(f"  Original: {original}")
    print(f"  Exported: {exported}")
    print("=" * 60)
    
    with zipfile.ZipFile(original, 'r') as zf_orig:
        with zipfile.ZipFile(exported, 'r') as zf_exp:
            orig_files = set(zf_orig.namelist())
            exp_files = set(zf_exp.namelist())
            
            # Check for missing/extra files
            missing = orig_files - exp_files
            extra = exp_files - orig_files
            
            if missing:
                print(f"\n❌ Missing in export: {missing}")
            if extra:
                print(f"\n⚠️  Extra in export: {extra}")
            
            # Compare common files
            identical = []
            different = []
            
            for name in sorted(orig_files & exp_files):
                orig_data = zf_orig.read(name)
                exp_data = zf_exp.read(name)
                
                if orig_data == exp_data:
                    identical.append(name)
                else:
                    different.append((name, len(orig_data), len(exp_data)))
            
            print(f"\n✓ Identical files: {len(identical)}")
            
            if different:
                print(f"\n⚠️  Different files ({len(different)}):")
                for name, orig_size, exp_size in different:
                    diff = exp_size - orig_size
                    print(f"  {name}: {orig_size} → {exp_size} ({diff:+d} bytes)")
            else:
                print("\n✓ All files identical!")


def main():
    if not TEST_FILE.exists():
        print(f"Test file not found: {TEST_FILE}")
        return
    
    print("Step 1: Parsing XLSX...")
    wb = xlsx_to_json(str(TEST_FILE), "debug-test")
    print(f"  Sheets: {len(wb.sheets)}")
    print(f"  Total cells: {sum(len(s.cells) for s in wb.sheets)}")
    
    # Mark a cell dirty to trigger export logic
    if wb.sheets and wb.sheets[0].cells:
        wb.sheets[0].cells[0].dirty = True
        print(f"  Marked cell {wb.sheets[0].cells[0].ref} as dirty")
    
    print("\nStep 2: Exporting XLSX...")
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    apply_json_to_xlsx(wb, str(TEST_FILE), str(OUTPUT_FILE))
    print(f"  Exported to: {OUTPUT_FILE}")
    
    print("\nStep 3: Comparing files...")
    compare_xlsx(TEST_FILE, OUTPUT_FILE)


if __name__ == "__main__":
    main()
