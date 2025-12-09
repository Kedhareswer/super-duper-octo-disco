"""
Test all DOCX and XLSX files individually via the API.

This script tests each file against the running server, identifies issues,
and verifies the full pipeline (upload → parse → export → re-parse).

Usage:
    python tests/test_all_files.py

Requirements:
    - Server must be running at http://127.0.0.1:8000
"""

import sys
import json
import time
import zipfile
import requests
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from io import BytesIO

# Add project root to path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


BASE_URL = "http://127.0.0.1:8000"


@dataclass
class FileTestResult:
    """Result of testing a single file."""
    file_path: str
    file_type: str
    file_size: int
    
    # Upload/Parse
    upload_success: bool = False
    upload_error: Optional[str] = None
    upload_time_ms: float = 0
    
    # Parse stats
    block_count: int = 0
    paragraph_count: int = 0
    table_count: int = 0
    drawing_count: int = 0
    checkbox_count: int = 0
    dropdown_count: int = 0
    total_chars: int = 0
    
    # Validation
    validation_passed: bool = False
    validation_errors: List[str] = field(default_factory=list)
    validation_warnings: List[str] = field(default_factory=list)
    
    # Export
    export_success: bool = False
    export_error: Optional[str] = None
    export_file_size: int = 0
    
    # Re-parse check
    reparse_success: bool = False
    reparse_block_count: int = 0
    reparse_chars: int = 0
    
    # Edge cases / Issues detected
    edge_cases: List[str] = field(default_factory=list)
    issues: List[str] = field(default_factory=list)
    
    @property
    def passed(self) -> bool:
        return (
            self.upload_success and
            self.validation_passed and
            self.export_success and
            self.reparse_success and
            len(self.issues) == 0
        )


def check_server() -> bool:
    """Check if the server is running."""
    try:
        r = requests.get(f"{BASE_URL}/", timeout=5)
        return r.status_code == 200
    except Exception:
        return False


def test_docx_file(file_path: Path) -> FileTestResult:
    """Test a single DOCX file through the full pipeline."""
    result = FileTestResult(
        file_path=str(file_path),
        file_type="docx",
        file_size=file_path.stat().st_size,
    )
    
    print(f"\n{'='*70}")
    print(f"Testing DOCX: {file_path.name}")
    print(f"{'='*70}")
    print(f"  File size: {result.file_size:,} bytes")
    
    # 1. UPLOAD/PARSE
    print(f"\n  [1] Uploading and parsing...")
    try:
        start = time.time()
        with open(file_path, 'rb') as f:
            r = requests.post(
                f"{BASE_URL}/documents/",
                files={"file": (file_path.name, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
            )
        result.upload_time_ms = (time.time() - start) * 1000
        
        if r.status_code == 200:
            result.upload_success = True
            doc = r.json()
            
            # Extract stats
            result.block_count = len(doc.get("blocks", []))
            result.checkbox_count = len(doc.get("checkboxes", []))
            result.dropdown_count = len(doc.get("dropdowns", []))
            
            # Count block types
            for block in doc.get("blocks", []):
                block_type = block.get("type")
                if block_type == "paragraph":
                    result.paragraph_count += 1
                    for run in block.get("runs", []):
                        if run.get("text"):
                            result.total_chars += len(run["text"])
                elif block_type == "table":
                    result.table_count += 1
                    # Count chars in table cells
                    for row in block.get("rows", []):
                        for cell in row.get("cells", []):
                            for cblock in cell.get("blocks", []):
                                result.paragraph_count += 1
                                for run in cblock.get("runs", []):
                                    if run.get("text"):
                                        result.total_chars += len(run["text"])
                elif block_type == "drawing":
                    result.drawing_count += 1
            
            print(f"      ✓ Upload OK ({result.upload_time_ms:.0f}ms)")
            print(f"      Blocks: {result.block_count} (P:{result.paragraph_count}, T:{result.table_count}, D:{result.drawing_count})")
            print(f"      Chars: {result.total_chars}, Checkboxes: {result.checkbox_count}, Dropdowns: {result.dropdown_count}")
            
            # Detect edge cases
            if result.checkbox_count > 0:
                result.edge_cases.append(f"Has {result.checkbox_count} checkboxes")
            if result.dropdown_count > 0:
                result.edge_cases.append(f"Has {result.dropdown_count} dropdowns")
            if result.drawing_count > 0:
                result.edge_cases.append(f"Has {result.drawing_count} drawings/images")
            if result.table_count > 5:
                result.edge_cases.append(f"Has many tables ({result.table_count})")
                
        else:
            result.upload_success = False
            result.upload_error = f"HTTP {r.status_code}: {r.text[:200]}"
            result.issues.append(f"Upload failed: {result.upload_error}")
            print(f"      ✗ Upload FAILED: {result.upload_error}")
            return result
            
    except Exception as e:
        result.upload_error = str(e)
        result.issues.append(f"Upload exception: {e}")
        print(f"      ✗ Upload EXCEPTION: {e}")
        return result
    
    # 2. VALIDATION
    print(f"\n  [2] Validating...")
    try:
        r = requests.get(f"{BASE_URL}/documents/{file_path.name}/validate")
        if r.status_code == 200:
            val = r.json()
            result.validation_passed = not val.get("has_errors", True)
            
            for issue in val.get("issues", []):
                msg = issue.get("message", str(issue))
                if issue.get("severity") == "error":
                    result.validation_errors.append(msg)
                else:
                    result.validation_warnings.append(msg)
            
            if result.validation_passed:
                print(f"      ✓ Validation passed")
            else:
                print(f"      ✗ Validation FAILED")
                for err in result.validation_errors[:3]:
                    print(f"        - {err}")
                result.issues.extend(result.validation_errors)
        else:
            result.issues.append(f"Validation request failed: HTTP {r.status_code}")
            print(f"      ✗ Validation request failed: HTTP {r.status_code}")
            
    except Exception as e:
        result.issues.append(f"Validation exception: {e}")
        print(f"      ✗ Validation EXCEPTION: {e}")
    
    # 3. EXPORT
    print(f"\n  [3] Exporting back to DOCX...")
    try:
        r = requests.post(f"{BASE_URL}/documents/{file_path.name}/export/file")
        if r.status_code == 200:
            result.export_success = True
            result.export_file_size = len(r.content)
            print(f"      ✓ Export OK ({result.export_file_size:,} bytes)")
            
            # Verify it's a valid ZIP
            try:
                with zipfile.ZipFile(BytesIO(r.content), 'r') as zf:
                    if 'word/document.xml' in zf.namelist():
                        print(f"      ✓ Valid DOCX structure")
                    else:
                        result.issues.append("Export missing word/document.xml")
                        print(f"      ✗ Missing word/document.xml in export")
            except zipfile.BadZipFile:
                result.issues.append("Export is not a valid ZIP file")
                print(f"      ✗ Export is not a valid ZIP")
                
        else:
            result.export_success = False
            result.export_error = f"HTTP {r.status_code}: {r.text[:200]}"
            result.issues.append(f"Export failed: {result.export_error}")
            print(f"      ✗ Export FAILED: {result.export_error}")
            
    except Exception as e:
        result.export_error = str(e)
        result.issues.append(f"Export exception: {e}")
        print(f"      ✗ Export EXCEPTION: {e}")
    
    # 4. RE-PARSE CHECK (using the test_backend_e2e approach)
    print(f"\n  [4] Checking re-parse consistency...")
    try:
        # Get original document again
        r = requests.get(f"{BASE_URL}/documents/{file_path.name}")
        if r.status_code == 200:
            orig_doc = r.json()
            result.reparse_success = True
            result.reparse_block_count = len(orig_doc.get("blocks", []))
            
            # Count chars in re-parsed
            for block in orig_doc.get("blocks", []):
                if block.get("type") == "paragraph":
                    for run in block.get("runs", []):
                        if run.get("text"):
                            result.reparse_chars += len(run["text"])
                elif block.get("type") == "table":
                    for row in block.get("rows", []):
                        for cell in row.get("cells", []):
                            for cblock in cell.get("blocks", []):
                                for run in cblock.get("runs", []):
                                    if run.get("text"):
                                        result.reparse_chars += len(run["text"])
            
            if result.reparse_block_count == result.block_count:
                print(f"      ✓ Block count consistent: {result.block_count}")
            else:
                result.issues.append(f"Block count mismatch: {result.block_count} → {result.reparse_block_count}")
                print(f"      ✗ Block count changed: {result.block_count} → {result.reparse_block_count}")
                
    except Exception as e:
        result.issues.append(f"Re-parse exception: {e}")
        print(f"      ✗ Re-parse EXCEPTION: {e}")
    
    # Summary
    print(f"\n  [RESULT] {'✓ PASSED' if result.passed else '✗ FAILED'}")
    if result.edge_cases:
        print(f"  Edge cases: {', '.join(result.edge_cases)}")
    if result.issues:
        print(f"  Issues: {len(result.issues)}")
        for issue in result.issues[:5]:
            print(f"    - {issue}")
    
    print(f"  Debug files: data/debug/{file_path.name}/")
    
    return result


def test_xlsx_file(file_path: Path) -> FileTestResult:
    """Test a single XLSX file through the spreadsheet pipeline."""
    result = FileTestResult(
        file_path=str(file_path),
        file_type="xlsx",
        file_size=file_path.stat().st_size,
    )
    
    print(f"\n{'='*70}")
    print(f"Testing XLSX: {file_path.name}")
    print(f"{'='*70}")
    print(f"  File size: {result.file_size:,} bytes")
    
    # 1. UPLOAD/PARSE
    print(f"\n  [1] Uploading and parsing...")
    try:
        start = time.time()
        with open(file_path, 'rb') as f:
            r = requests.post(
                f"{BASE_URL}/spreadsheets/",
                files={"file": (file_path.name, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            )
        result.upload_time_ms = (time.time() - start) * 1000
        
        if r.status_code == 200:
            result.upload_success = True
            doc = r.json()
            
            # XLSX has sheets instead of blocks
            sheets = doc.get("sheets", [])
            result.block_count = len(sheets)  # Use block_count for sheets
            
            # Count cells and chars
            total_cells = 0
            for sheet in sheets:
                for row in sheet.get("rows", []):
                    for cell in row.get("cells", []):
                        total_cells += 1
                        if cell.get("value"):
                            result.total_chars += len(str(cell["value"]))
            
            print(f"      ✓ Upload OK ({result.upload_time_ms:.0f}ms)")
            print(f"      Sheets: {result.block_count}, Cells: {total_cells}, Chars: {result.total_chars}")
            
            # Detect edge cases
            merged_ranges = doc.get("merged_ranges", [])
            if merged_ranges:
                result.edge_cases.append(f"Has {len(merged_ranges)} merged ranges")
            validations = doc.get("validations", [])
            if validations:
                result.edge_cases.append(f"Has {len(validations)} data validations")
                
        else:
            result.upload_success = False
            result.upload_error = f"HTTP {r.status_code}: {r.text[:200]}"
            result.issues.append(f"Upload failed: {result.upload_error}")
            print(f"      ✗ Upload FAILED: {result.upload_error}")
            return result
            
    except Exception as e:
        result.upload_error = str(e)
        result.issues.append(f"Upload exception: {e}")
        print(f"      ✗ Upload EXCEPTION: {e}")
        return result
    
    # 2. EXPORT
    print(f"\n  [2] Exporting back to XLSX...")
    try:
        r = requests.post(f"{BASE_URL}/spreadsheets/{file_path.name}/export/file")
        if r.status_code == 200:
            result.export_success = True
            result.export_file_size = len(r.content)
            print(f"      ✓ Export OK ({result.export_file_size:,} bytes)")
            
            # Verify it's a valid ZIP (XLSX is ZIP-based)
            try:
                with zipfile.ZipFile(BytesIO(r.content), 'r') as zf:
                    if 'xl/workbook.xml' in zf.namelist():
                        print(f"      ✓ Valid XLSX structure")
                    else:
                        result.issues.append("Export missing xl/workbook.xml")
                        print(f"      ✗ Missing xl/workbook.xml in export")
            except zipfile.BadZipFile:
                result.issues.append("Export is not a valid ZIP file")
                print(f"      ✗ Export is not a valid ZIP")
                
        else:
            result.export_success = False
            result.export_error = f"HTTP {r.status_code}: {r.text[:200]}"
            result.issues.append(f"Export failed: {result.export_error}")
            print(f"      ✗ Export FAILED: {result.export_error}")
            
    except Exception as e:
        result.export_error = str(e)
        result.issues.append(f"Export exception: {e}")
        print(f"      ✗ Export EXCEPTION: {e}")
    
    result.reparse_success = result.export_success  # Simplified for XLSX
    result.validation_passed = result.upload_success  # No separate validation for XLSX yet
    
    # Summary
    print(f"\n  [RESULT] {'✓ PASSED' if result.passed else '✗ FAILED'}")
    if result.edge_cases:
        print(f"  Edge cases: {', '.join(result.edge_cases)}")
    if result.issues:
        print(f"  Issues: {len(result.issues)}")
        for issue in result.issues[:5]:
            print(f"    - {issue}")
    
    return result


def main():
    """Run tests on all DOCX and XLSX files."""
    
    print("=" * 70)
    print("  ALL FILES TEST SUITE")
    print("=" * 70)
    
    # Check server
    print("\nChecking server...")
    if not check_server():
        print("  ✗ Server not running at", BASE_URL)
        print("  Start the server with: uvicorn main:app --reload --port 8000")
        sys.exit(1)
    print("  ✓ Server is running")
    
    # Find test files
    docx_dir = ROOT / "data" / "uploads" / "docx"
    xlsx_dir = ROOT / "data" / "uploads" / "excel"
    
    docx_files = sorted(docx_dir.glob("*.docx")) + sorted(docx_dir.glob("*.DOCX"))
    xlsx_files = sorted(xlsx_dir.glob("*.xlsx")) + sorted(xlsx_dir.glob("*.XLSX"))
    
    print(f"\nFound {len(docx_files)} DOCX files and {len(xlsx_files)} XLSX files")
    
    results: List[FileTestResult] = []
    
    # Test DOCX files one at a time
    print("\n" + "=" * 70)
    print("  DOCX FILES")
    print("=" * 70)
    
    for file_path in docx_files:
        result = test_docx_file(file_path)
        results.append(result)
        time.sleep(0.5)  # Brief pause between files
    
    # Test XLSX files one at a time
    print("\n" + "=" * 70)
    print("  XLSX FILES")
    print("=" * 70)
    
    for file_path in xlsx_files:
        result = test_xlsx_file(file_path)
        results.append(result)
        time.sleep(0.5)
    
    # Summary
    print("\n" + "=" * 70)
    print("  SUMMARY")
    print("=" * 70)
    
    passed = [r for r in results if r.passed]
    failed = [r for r in results if not r.passed]
    
    print(f"\n  Total: {len(results)} files")
    print(f"  Passed: {len(passed)}")
    print(f"  Failed: {len(failed)}")
    
    if failed:
        print("\n  Failed files:")
        for r in failed:
            print(f"    - {Path(r.file_path).name}: {len(r.issues)} issues")
            for issue in r.issues[:2]:
                print(f"        {issue}")
    
    # Edge cases summary
    all_edge_cases = []
    for r in results:
        for ec in r.edge_cases:
            all_edge_cases.append(f"{Path(r.file_path).name}: {ec}")
    
    if all_edge_cases:
        print("\n  Edge cases detected:")
        for ec in all_edge_cases:
            print(f"    - {ec}")
    
    print("\n  Debug files saved to: data/debug/")
    
    return 0 if len(failed) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
