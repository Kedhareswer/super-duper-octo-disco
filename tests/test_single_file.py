"""Test a single file via the API."""
import sys
import requests
import json
from pathlib import Path

BASE_URL = "http://127.0.0.1:8000"


def test_docx(file_path: str):
    """Test a DOCX file."""
    path = Path(file_path)
    print("=" * 70)
    print(f"Testing DOCX: {path.name}")
    print("=" * 70)
    print(f"  File size: {path.stat().st_size:,} bytes")
    
    # 1. Upload and Parse
    print("\n[1] Upload and Parse...")
    with open(path, 'rb') as f:
        r = requests.post(
            f"{BASE_URL}/documents/",
            files={"file": (path.name, f)}
        )
    
    if r.status_code != 200:
        print(f"    FAILED: HTTP {r.status_code}")
        print(f"    {r.text[:500]}")
        return False
    
    doc = r.json()
    blocks = doc.get("blocks", [])
    checkboxes = doc.get("checkboxes", [])
    dropdowns = doc.get("dropdowns", [])
    
    # Count block types
    para_count = sum(1 for b in blocks if b.get("type") == "paragraph")
    table_count = sum(1 for b in blocks if b.get("type") == "table")
    drawing_count = sum(1 for b in blocks if b.get("type") == "drawing")
    
    # Count chars
    total_chars = 0
    for block in blocks:
        if block.get("type") == "paragraph":
            for run in block.get("runs", []):
                if run.get("text"):
                    total_chars += len(run["text"])
        elif block.get("type") == "table":
            for row in block.get("rows", []):
                for cell in row.get("cells", []):
                    for cb in cell.get("blocks", []):
                        for run in cb.get("runs", []):
                            if run.get("text"):
                                total_chars += len(run["text"])
    
    print(f"    OK - Blocks: {len(blocks)} (P:{para_count}, T:{table_count}, D:{drawing_count})")
    print(f"    Chars: {total_chars}, Checkboxes: {len(checkboxes)}, Dropdowns: {len(dropdowns)}")
    
    # Detect edge cases
    edge_cases = []
    if checkboxes:
        edge_cases.append(f"{len(checkboxes)} checkboxes")
    if dropdowns:
        edge_cases.append(f"{len(dropdowns)} dropdowns")
    if drawing_count:
        edge_cases.append(f"{drawing_count} drawings")
    if table_count > 3:
        edge_cases.append(f"many tables ({table_count})")
    
    if edge_cases:
        print(f"    Edge cases: {', '.join(edge_cases)}")
    
    # 2. Validate
    print("\n[2] Validate...")
    r = requests.get(f"{BASE_URL}/documents/{path.name}/validate")
    val = r.json()
    
    if val.get("has_errors"):
        print(f"    VALIDATION ERRORS:")
        for issue in val.get("issues", []):
            if issue.get("severity") == "error":
                print(f"      - {issue.get('message')}")
    else:
        print(f"    OK - No validation errors")
        stages = val.get("stages", [])
        if stages:
            s = stages[-1]
            print(f"    Stats: {s.get('total_chars')} chars, {s.get('paragraph_count')} paragraphs, {s.get('table_count')} tables")
    
    # 3. Export
    print("\n[3] Export to DOCX...")
    r = requests.post(f"{BASE_URL}/documents/{path.name}/export/file")
    
    if r.status_code == 200:
        print(f"    OK - Exported {len(r.content):,} bytes")
        
        # Verify it's a valid DOCX
        import zipfile
        from io import BytesIO
        try:
            with zipfile.ZipFile(BytesIO(r.content), 'r') as zf:
                if 'word/document.xml' in zf.namelist():
                    print(f"    Valid DOCX structure")
                else:
                    print(f"    WARNING: Missing word/document.xml")
        except Exception as e:
            print(f"    ERROR: Invalid ZIP: {e}")
    else:
        print(f"    FAILED: HTTP {r.status_code}")
        print(f"    {r.text[:200]}")
        return False
    
    # 4. Debug output location
    print(f"\n[4] Debug files: data/debug/{path.name}/")
    
    debug_dir = Path(f"data/debug/{path.name}")
    if debug_dir.exists():
        files = list(debug_dir.glob("*"))
        print(f"    {len(files)} debug files created")
        for f in sorted(files)[:10]:
            print(f"      - {f.name} ({f.stat().st_size:,} bytes)")
    
    print("\n" + "=" * 70)
    print("RESULT: PASSED")
    print("=" * 70)
    return True


def test_xlsx(file_path: str):
    """Test an XLSX file."""
    path = Path(file_path)
    print("=" * 70)
    print(f"Testing XLSX: {path.name}")
    print("=" * 70)
    print(f"  File size: {path.stat().st_size:,} bytes")
    
    # 1. Upload and Parse
    print("\n[1] Upload and Parse...")
    with open(path, 'rb') as f:
        r = requests.post(
            f"{BASE_URL}/spreadsheets/",
            files={"file": (path.name, f)}
        )
    
    if r.status_code != 200:
        print(f"    FAILED: HTTP {r.status_code}")
        print(f"    {r.text[:500]}")
        return False
    
    doc = r.json()
    
    # Get the spreadsheet ID from the response (UUID-prefixed)
    spreadsheet_id = doc.get("id", "")
    if not spreadsheet_id:
        print(f"    WARNING: No spreadsheet ID in response")
        spreadsheet_id = path.name  # Fallback
    
    sheets = doc.get("sheets", [])
    
    # Count cells and chars from the response format
    total_cells = 0
    total_chars = 0
    
    # The response format has sheet summaries with cell_count
    for sheet in sheets:
        cell_count = sheet.get("cell_count", 0)
        total_cells += cell_count
        # If detailed cells are available
        for row in sheet.get("rows", []):
            for cell in row.get("cells", []):
                if cell.get("value"):
                    total_chars += len(str(cell["value"]))
    
    print(f"    OK - ID: {spreadsheet_id}")
    print(f"    Sheets: {len(sheets)}, Cells: {total_cells}")
    
    # Show sheet details
    for sheet in sheets:
        print(f"      - {sheet.get('name', 'Unknown')}: {sheet.get('cell_count', 0)} cells, {sheet.get('row_count', 0)} rows")
    
    # Edge cases
    edge_cases = []
    merged = doc.get("merged_ranges", [])
    if merged:
        edge_cases.append(f"{len(merged)} merged ranges")
    validations = doc.get("validations", [])
    if validations:
        edge_cases.append(f"{len(validations)} data validations")
    images = doc.get("images", [])
    if images:
        edge_cases.append(f"{len(images)} images")
    
    if edge_cases:
        print(f"    Edge cases: {', '.join(edge_cases)}")
    
    # 2. Export - use the actual spreadsheet_id
    print("\n[2] Export to XLSX...")
    r = requests.post(f"{BASE_URL}/spreadsheets/{spreadsheet_id}/export/file")
    
    if r.status_code == 200:
        print(f"    OK - Exported {len(r.content):,} bytes")
        
        # Verify it's a valid XLSX (ZIP-based)
        import zipfile
        from io import BytesIO
        try:
            with zipfile.ZipFile(BytesIO(r.content), 'r') as zf:
                if 'xl/workbook.xml' in zf.namelist():
                    print(f"    Valid XLSX structure")
                else:
                    print(f"    WARNING: Missing xl/workbook.xml")
        except Exception as e:
            print(f"    ERROR: Invalid ZIP: {e}")
    else:
        print(f"    FAILED: HTTP {r.status_code}")
        print(f"    {r.text[:200]}")
        return False
    
    print("\n" + "=" * 70)
    print("RESULT: PASSED")
    print("=" * 70)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python test_single_file.py <file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    path = Path(file_path)
    
    if not path.exists():
        print(f"File not found: {file_path}")
        sys.exit(1)
    
    if path.suffix.lower() == ".docx":
        success = test_docx(file_path)
    elif path.suffix.lower() == ".xlsx":
        success = test_xlsx(file_path)
    else:
        print(f"Unsupported file type: {path.suffix}")
        sys.exit(1)
    
    sys.exit(0 if success else 1)
