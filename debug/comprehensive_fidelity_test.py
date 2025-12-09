"""
Comprehensive Fidelity Test for ALL DOCX and XLSX files.
==========================================================

Tests:
1. Parsing - Does it extract all content?
2. Roundtrip - Does export preserve content?
3. Fidelity - No loss of text, structure, formatting?
4. Export - Can Microsoft Word/Excel open the files?
"""

import os
import sys
import zipfile
import hashlib
import traceback
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from collections import Counter
from io import BytesIO
from xml.etree import ElementTree as ET

# Add project root to path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from services.document_engine import docx_to_json, apply_json_to_docx
from services.validation import (
    extract_raw_docx_content,
    extract_json_content,
    ValidationReport,
    ValidationIssue,
    StageSnapshot,
)
from models.schemas import ParagraphBlock, TableBlock, DrawingBlock

from services.excel_engine import xlsx_to_json, apply_json_to_xlsx


# =============================================================================
# DATA CLASSES
# =============================================================================

@dataclass
class FileTestResult:
    """Result of testing a single file."""
    file_path: str
    file_type: str  # "docx" or "xlsx"
    file_size: int
    
    # Parse results
    parse_success: bool = False
    parse_error: Optional[str] = None
    parse_time_ms: float = 0
    
    # Content stats (after parse)
    block_count: int = 0
    paragraph_count: int = 0
    table_count: int = 0
    drawing_count: int = 0
    total_chars: int = 0
    run_count: int = 0
    checkbox_count: int = 0
    dropdown_count: int = 0
    
    # For XLSX
    sheet_count: int = 0
    cell_count: int = 0
    merged_range_count: int = 0
    validation_count: int = 0
    image_count: int = 0
    comment_count: int = 0
    
    # Roundtrip results
    export_success: bool = False
    export_error: Optional[str] = None
    export_file_size: int = 0
    
    # Fidelity results
    char_loss: int = 0
    block_loss: int = 0
    structure_issues: List[str] = field(default_factory=list)
    
    # Re-parse results (of exported file)
    reparse_success: bool = False
    reparse_error: Optional[str] = None
    reparse_block_count: int = 0
    reparse_chars: int = 0
    
    # Issues
    issues: List[str] = field(default_factory=list)
    
    @property
    def passed(self) -> bool:
        return (
            self.parse_success and 
            self.export_success and 
            self.reparse_success and 
            self.char_loss == 0 and 
            self.block_loss == 0 and
            len(self.issues) == 0
        )


# =============================================================================
# DOCX TESTING
# =============================================================================

def test_docx_file(file_path: Path) -> FileTestResult:
    """Comprehensive test of a single DOCX file."""
    result = FileTestResult(
        file_path=str(file_path),
        file_type="docx",
        file_size=file_path.stat().st_size,
    )
    
    # 1. PARSE TEST
    try:
        import time
        start = time.time()
        json_doc = docx_to_json(str(file_path), file_path.name)
        result.parse_time_ms = (time.time() - start) * 1000
        result.parse_success = True
        
        # Count content
        result.block_count = len(json_doc.blocks)
        result.checkbox_count = len(json_doc.checkboxes)
        result.dropdown_count = len(json_doc.dropdowns)
        
        for block in json_doc.blocks:
            if isinstance(block, ParagraphBlock):
                result.paragraph_count += 1
                for run in block.runs:
                    result.run_count += 1
                    if run.text:
                        result.total_chars += len(run.text)
            elif isinstance(block, TableBlock):
                result.table_count += 1
                for row in block.rows:
                    for cell in row.cells:
                        for para in cell.blocks:
                            result.paragraph_count += 1
                            for run in para.runs:
                                result.run_count += 1
                                if run.text:
                                    result.total_chars += len(run.text)
            elif isinstance(block, DrawingBlock):
                result.drawing_count += 1
        
    except Exception as e:
        result.parse_success = False
        result.parse_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"PARSE ERROR: {result.parse_error}")
        return result
    
    # 2. EXPORT TEST (roundtrip)
    output_dir = ROOT / "data" / "test_outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"roundtrip_{file_path.name}"
    
    try:
        apply_json_to_docx(json_doc, str(file_path), str(output_path))
        result.export_success = True
        result.export_file_size = output_path.stat().st_size
        
        # Verify exported file is valid DOCX
        with zipfile.ZipFile(output_path, 'r') as zf:
            if 'word/document.xml' not in zf.namelist():
                result.issues.append("EXPORT: Missing word/document.xml")
            
            # Check for bad namespace prefixes
            content = zf.read('word/document.xml')
            import re
            bad_ns = re.findall(rb'<ns\d+:', content)
            if bad_ns:
                result.issues.append(f"EXPORT: Found {len(bad_ns)} bad namespace prefixes (ns0:, ns1:)")
            
            # Check XML validity
            try:
                ET.parse(BytesIO(content))
            except ET.ParseError as e:
                result.issues.append(f"EXPORT: Invalid XML - {e}")
        
    except Exception as e:
        result.export_success = False
        result.export_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"EXPORT ERROR: {result.export_error}")
        return result
    
    # 3. RE-PARSE TEST (fidelity)
    try:
        reparsed = docx_to_json(str(output_path), "reparsed")
        result.reparse_success = True
        result.reparse_block_count = len(reparsed.blocks)
        
        # Count chars in reparsed
        for block in reparsed.blocks:
            if isinstance(block, ParagraphBlock):
                for run in block.runs:
                    if run.text:
                        result.reparse_chars += len(run.text)
            elif isinstance(block, TableBlock):
                for row in block.rows:
                    for cell in row.cells:
                        for para in cell.blocks:
                            for run in para.runs:
                                if run.text:
                                    result.reparse_chars += len(run.text)
        
        # Calculate loss
        result.char_loss = result.total_chars - result.reparse_chars
        result.block_loss = result.block_count - result.reparse_block_count
        
        if result.char_loss > 0:
            result.issues.append(f"FIDELITY: Lost {result.char_loss} characters")
        elif result.char_loss < 0:
            result.issues.append(f"FIDELITY: Gained {abs(result.char_loss)} characters (possible duplicate)")
        
        if result.block_loss > 0:
            result.issues.append(f"FIDELITY: Lost {result.block_loss} blocks")
        elif result.block_loss < 0:
            result.issues.append(f"FIDELITY: Gained {abs(result.block_loss)} blocks")
        
    except Exception as e:
        result.reparse_success = False
        result.reparse_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"REPARSE ERROR: {result.reparse_error}")
    
    return result


# =============================================================================
# XLSX TESTING
# =============================================================================

def test_xlsx_file(file_path: Path) -> FileTestResult:
    """Comprehensive test of a single XLSX file."""
    result = FileTestResult(
        file_path=str(file_path),
        file_type="xlsx",
        file_size=file_path.stat().st_size,
    )
    
    # 1. PARSE TEST
    try:
        import time
        start = time.time()
        workbook = xlsx_to_json(str(file_path), file_path.name)
        result.parse_time_ms = (time.time() - start) * 1000
        result.parse_success = True
        
        # Count content
        result.sheet_count = len(workbook.sheets)
        
        for sheet in workbook.sheets:
            result.cell_count += len(sheet.cells)
            result.merged_range_count += len(sheet.merged_cells)
            result.validation_count += len(sheet.data_validations)
            result.image_count += len(sheet.images)
            result.comment_count += len(sheet.comments)
            
            # Count chars
            for cell in sheet.cells:
                if cell.value and isinstance(cell.value, str):
                    result.total_chars += len(cell.value)
        
    except Exception as e:
        result.parse_success = False
        result.parse_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"PARSE ERROR: {result.parse_error}")
        return result
    
    # 2. EXPORT TEST (roundtrip)
    output_dir = ROOT / "data" / "test_outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"roundtrip_{file_path.name}"
    
    try:
        # Mark first cell as dirty to trigger export
        if workbook.sheets and workbook.sheets[0].cells:
            first_cell = workbook.sheets[0].cells[0]
            first_cell.original_value = first_cell.value
            first_cell.dirty = True
        
        apply_json_to_xlsx(workbook, str(file_path), str(output_path))
        result.export_success = True
        result.export_file_size = output_path.stat().st_size
        
        # Verify exported file is valid XLSX
        with zipfile.ZipFile(output_path, 'r') as zf:
            if 'xl/workbook.xml' not in zf.namelist():
                result.issues.append("EXPORT: Missing xl/workbook.xml")
            
            # Check XML validity for each sheet
            for name in zf.namelist():
                if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                    try:
                        content = zf.read(name)
                        ET.parse(BytesIO(content))
                    except ET.ParseError as e:
                        result.issues.append(f"EXPORT: Invalid XML in {name} - {e}")
        
    except Exception as e:
        result.export_success = False
        result.export_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"EXPORT ERROR: {result.export_error}")
        return result
    
    # 3. RE-PARSE TEST (fidelity)
    try:
        reparsed = xlsx_to_json(str(output_path), "reparsed")
        result.reparse_success = True
        result.reparse_block_count = len(reparsed.sheets)  # Use sheet count for xlsx
        
        # Count cells and chars in reparsed
        reparse_cells = 0
        for sheet in reparsed.sheets:
            reparse_cells += len(sheet.cells)
            for cell in sheet.cells:
                if cell.value and isinstance(cell.value, str):
                    result.reparse_chars += len(cell.value)
        
        # Calculate loss
        result.char_loss = result.total_chars - result.reparse_chars
        result.block_loss = result.cell_count - reparse_cells
        
        if result.char_loss > 0:
            result.issues.append(f"FIDELITY: Lost {result.char_loss} characters")
        
        if result.block_loss > 0:
            result.issues.append(f"FIDELITY: Lost {result.block_loss} cells")
        
        # Check merged cells preserved
        orig_merges = sum(len(s.merged_cells) for s in workbook.sheets)
        new_merges = sum(len(s.merged_cells) for s in reparsed.sheets)
        if new_merges < orig_merges:
            result.issues.append(f"FIDELITY: Lost {orig_merges - new_merges} merged cell ranges")
        
    except Exception as e:
        result.reparse_success = False
        result.reparse_error = f"{type(e).__name__}: {str(e)}"
        result.issues.append(f"REPARSE ERROR: {result.reparse_error}")
    
    return result


# =============================================================================
# MAIN TEST RUNNER
# =============================================================================

def run_all_tests():
    """Run comprehensive tests on all DOCX and XLSX files."""
    
    print("="*80)
    print("COMPREHENSIVE FIDELITY TEST")
    print("="*80)
    
    # Find all test files
    docx_dir = ROOT / "data" / "uploads" / "docx"
    xlsx_dir = ROOT / "data" / "uploads" / "excel"
    
    docx_files = list(docx_dir.glob("*.docx")) + list(docx_dir.glob("*.DOCX"))
    xlsx_files = list(xlsx_dir.glob("*.xlsx")) + list(xlsx_dir.glob("*.XLSX"))
    
    print(f"\nFound {len(docx_files)} DOCX files and {len(xlsx_files)} XLSX files\n")
    
    all_results: List[FileTestResult] = []
    
    # Test DOCX files
    print("\n" + "="*80)
    print("DOCX FILES")
    print("="*80)
    
    for file_path in sorted(docx_files):
        print(f"\n--- Testing: {file_path.name} ({file_path.stat().st_size:,} bytes) ---")
        result = test_docx_file(file_path)
        all_results.append(result)
        
        # Print summary
        if result.parse_success:
            print(f"  ✓ Parse: {result.block_count} blocks, {result.total_chars:,} chars, {result.parse_time_ms:.0f}ms")
        else:
            print(f"  ✗ Parse FAILED: {result.parse_error}")
        
        if result.export_success:
            print(f"  ✓ Export: {result.export_file_size:,} bytes")
        elif result.parse_success:
            print(f"  ✗ Export FAILED: {result.export_error}")
        
        if result.reparse_success:
            status = "✓" if result.char_loss == 0 and result.block_loss == 0 else "⚠"
            print(f"  {status} Fidelity: char_loss={result.char_loss}, block_loss={result.block_loss}")
        elif result.export_success:
            print(f"  ✗ Reparse FAILED: {result.reparse_error}")
        
        if result.issues:
            print(f"  Issues ({len(result.issues)}):")
            for issue in result.issues:
                print(f"    - {issue}")
    
    # Test XLSX files
    print("\n" + "="*80)
    print("XLSX FILES")
    print("="*80)
    
    for file_path in sorted(xlsx_files):
        print(f"\n--- Testing: {file_path.name} ({file_path.stat().st_size:,} bytes) ---")
        result = test_xlsx_file(file_path)
        all_results.append(result)
        
        # Print summary
        if result.parse_success:
            print(f"  ✓ Parse: {result.sheet_count} sheets, {result.cell_count:,} cells, {result.parse_time_ms:.0f}ms")
        else:
            print(f"  ✗ Parse FAILED: {result.parse_error}")
        
        if result.export_success:
            print(f"  ✓ Export: {result.export_file_size:,} bytes")
        elif result.parse_success:
            print(f"  ✗ Export FAILED: {result.export_error}")
        
        if result.reparse_success:
            status = "✓" if result.char_loss == 0 and result.block_loss == 0 else "⚠"
            print(f"  {status} Fidelity: char_loss={result.char_loss}, cell_loss={result.block_loss}")
        elif result.export_success:
            print(f"  ✗ Reparse FAILED: {result.reparse_error}")
        
        if result.issues:
            print(f"  Issues ({len(result.issues)}):")
            for issue in result.issues:
                print(f"    - {issue}")
    
    # Summary
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    
    docx_results = [r for r in all_results if r.file_type == "docx"]
    xlsx_results = [r for r in all_results if r.file_type == "xlsx"]
    
    print(f"\nDOCX Files ({len(docx_results)} total):")
    print(f"  Parse:  {sum(1 for r in docx_results if r.parse_success)}/{len(docx_results)} passed")
    print(f"  Export: {sum(1 for r in docx_results if r.export_success)}/{len(docx_results)} passed")
    print(f"  Reparse: {sum(1 for r in docx_results if r.reparse_success)}/{len(docx_results)} passed")
    print(f"  Perfect fidelity: {sum(1 for r in docx_results if r.passed)}/{len(docx_results)}")
    
    print(f"\nXLSX Files ({len(xlsx_results)} total):")
    print(f"  Parse:  {sum(1 for r in xlsx_results if r.parse_success)}/{len(xlsx_results)} passed")
    print(f"  Export: {sum(1 for r in xlsx_results if r.export_success)}/{len(xlsx_results)} passed")
    print(f"  Reparse: {sum(1 for r in xlsx_results if r.reparse_success)}/{len(xlsx_results)} passed")
    print(f"  Perfect fidelity: {sum(1 for r in xlsx_results if r.passed)}/{len(xlsx_results)}")
    
    # List all issues
    all_issues = []
    for r in all_results:
        for issue in r.issues:
            all_issues.append((Path(r.file_path).name, issue))
    
    if all_issues:
        print(f"\n" + "="*80)
        print(f"ALL ISSUES ({len(all_issues)} total)")
        print("="*80)
        
        for filename, issue in all_issues:
            print(f"  [{filename}] {issue}")
    
    # Recommendations
    print(f"\n" + "="*80)
    print("RECOMMENDATIONS")
    print("="*80)
    
    # Collect patterns
    parse_errors = [r for r in all_results if not r.parse_success]
    export_errors = [r for r in all_results if r.parse_success and not r.export_success]
    fidelity_issues = [r for r in all_results if r.reparse_success and (r.char_loss != 0 or r.block_loss != 0)]
    
    if parse_errors:
        print(f"\n1. PARSE ISSUES ({len(parse_errors)} files):")
        for r in parse_errors:
            print(f"   - {Path(r.file_path).name}: {r.parse_error}")
    
    if export_errors:
        print(f"\n2. EXPORT ISSUES ({len(export_errors)} files):")
        for r in export_errors:
            print(f"   - {Path(r.file_path).name}: {r.export_error}")
    
    if fidelity_issues:
        print(f"\n3. FIDELITY ISSUES ({len(fidelity_issues)} files):")
        for r in fidelity_issues:
            print(f"   - {Path(r.file_path).name}: char_loss={r.char_loss}, block_loss={r.block_loss}")
    
    if not parse_errors and not export_errors and not fidelity_issues:
        print("\n✅ ALL FILES PASSED WITH PERFECT FIDELITY!")
    
    return all_results


if __name__ == "__main__":
    results = run_all_tests()
    
    # Return exit code
    failed = sum(1 for r in results if not r.passed)
    sys.exit(0 if failed == 0 else 1)
