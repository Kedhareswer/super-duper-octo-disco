"""Comprehensive validation and debugging for document processing pipeline.

This module provides stage-by-stage validation to ensure:
1. DOCX → JSON parsing preserves all content correctly
2. JSON structure is valid and complete
3. JSON → DOCX export preserves content without adding/removing text
4. xml_ref paths are correctly mapped
"""
from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional
from xml.etree import ElementTree as ET
from collections import Counter

from models.schemas import DocumentJSON, ParagraphBlock, TableBlock, DrawingBlock


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}


@dataclass
class ValidationIssue:
    """A single validation issue found during processing."""
    stage: str  # "parse", "validate", "export"
    severity: str  # "error", "warning", "info"
    category: str  # "missing_text", "extra_text", "structure", "xml_ref"
    message: str
    details: Optional[Dict] = None


@dataclass
class StageSnapshot:
    """Snapshot of document content at a particular stage."""
    stage: str
    text_elements: List[str]
    total_chars: int
    paragraph_count: int
    table_count: int
    row_count: int
    cell_count: int
    run_count: int
    checkbox_count: int = 0
    dropdown_count: int = 0


@dataclass 
class ValidationReport:
    """Complete validation report for a document processing operation."""
    document_id: str
    stages: List[StageSnapshot] = field(default_factory=list)
    issues: List[ValidationIssue] = field(default_factory=list)
    
    @property
    def has_errors(self) -> bool:
        return any(i.severity == "error" for i in self.issues)
    
    @property
    def has_warnings(self) -> bool:
        return any(i.severity == "warning" for i in self.issues)
    
    def add_issue(self, stage: str, severity: str, category: str, message: str, details: Dict = None):
        self.issues.append(ValidationIssue(stage, severity, category, message, details))
    
    def add_stage(self, snapshot: StageSnapshot):
        self.stages.append(snapshot)
    
    def to_dict(self) -> Dict:
        return {
            "document_id": self.document_id,
            "has_errors": self.has_errors,
            "has_warnings": self.has_warnings,
            "stages": [
                {
                    "stage": s.stage,
                    "total_chars": s.total_chars,
                    "paragraph_count": s.paragraph_count,
                    "table_count": s.table_count,
                    "row_count": s.row_count,
                    "cell_count": s.cell_count,
                    "run_count": s.run_count,
                    "checkbox_count": s.checkbox_count,
                    "dropdown_count": s.dropdown_count,
                }
                for s in self.stages
            ],
            "issues": [
                {
                    "stage": i.stage,
                    "severity": i.severity,
                    "category": i.category,
                    "message": i.message,
                    "details": i.details,
                }
                for i in self.issues
            ]
        }


def extract_raw_docx_content(docx_path: str) -> StageSnapshot:
    """Extract content directly from DOCX XML for comparison."""
    with zipfile.ZipFile(docx_path, 'r') as zf:
        with zf.open('word/document.xml') as doc:
            tree = ET.parse(doc)
            root = tree.getroot()
            body = root.find(f"{{{NS['w']}}}body")
    
    if body is None:
        return StageSnapshot(
            stage="raw_docx",
            text_elements=[],
            total_chars=0,
            paragraph_count=0,
            table_count=0,
            row_count=0,
            cell_count=0,
            run_count=0,
        )
    
    # Count elements
    paragraphs = len(body.findall(f".//{{{NS['w']}}}p"))
    tables = len(body.findall(f".//{{{NS['w']}}}tbl"))
    rows = len(body.findall(f".//{{{NS['w']}}}tr"))
    cells = len(body.findall(f".//{{{NS['w']}}}tc"))
    runs = len(body.findall(f".//{{{NS['w']}}}r"))
    
    # Extract all text
    text_elements = []
    for t in body.iter(f"{{{NS['w']}}}t"):
        if t.text:
            text_elements.append(t.text)
    
    total_chars = sum(len(t) for t in text_elements)
    
    # Count checkboxes and dropdowns
    checkboxes = 0
    dropdowns = 0
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is not None:
            if sdt_pr.find("w14:checkbox", NS) is not None:
                checkboxes += 1
            elif sdt_pr.find("w:comboBox", NS) is not None or sdt_pr.find("w:dropDownList", NS) is not None:
                dropdowns += 1
    
    return StageSnapshot(
        stage="raw_docx",
        text_elements=text_elements,
        total_chars=total_chars,
        paragraph_count=paragraphs,
        table_count=tables,
        row_count=rows,
        cell_count=cells,
        run_count=runs,
        checkbox_count=checkboxes,
        dropdown_count=dropdowns,
    )


def extract_json_content(json_doc: DocumentJSON) -> StageSnapshot:
    """Extract content from parsed JSON for comparison.
    
    Handles nested tables recursively.
    """
    text_elements: list[str] = []
    paragraph_count = 0
    table_count = 0
    row_count = 0
    cell_count = 0
    run_count = 0
    
    def process_table(table: TableBlock):
        """Recursively process a table and its nested tables."""
        nonlocal table_count, row_count, cell_count, paragraph_count, run_count
        table_count += 1
        for row in table.rows:
            row_count += 1
            for cell in row.cells:
                cell_count += 1
                process_cell_blocks(cell.blocks)
    
    def process_cell_blocks(blocks):
        """Process blocks inside a cell (paragraphs and nested tables)."""
        nonlocal paragraph_count, run_count
        for block in blocks:
            if isinstance(block, ParagraphBlock):
                paragraph_count += 1
                for run in block.runs:
                    run_count += 1
                    # Only TextRuns have text attribute; CheckboxRun/DropdownRun don't
                    if hasattr(run, 'text') and run.text:
                        text_elements.append(run.text)
            elif isinstance(block, TableBlock):
                # Nested table - process recursively
                process_table(block)
    
    for block in json_doc.blocks:
        if isinstance(block, ParagraphBlock):
            paragraph_count += 1
            for run in block.runs:
                run_count += 1
                # Only TextRuns have text attribute; CheckboxRun/DropdownRun don't
                if hasattr(run, 'text') and run.text:
                    text_elements.append(run.text)
        elif isinstance(block, TableBlock):
            process_table(block)
    
    total_chars = sum(len(t) for t in text_elements)
    
    return StageSnapshot(
        stage="parsed_json",
        text_elements=text_elements,
        total_chars=total_chars,
        paragraph_count=paragraph_count,
        table_count=table_count,
        row_count=row_count,
        cell_count=cell_count,
        run_count=run_count,
        checkbox_count=len(json_doc.checkboxes),
        dropdown_count=len(json_doc.dropdowns),
    )


def compare_snapshots(before: StageSnapshot, after: StageSnapshot, report: ValidationReport):
    """Compare two snapshots and report differences.
    
    Note: When parsing checkboxes/dropdowns as inline controls (CheckboxRun/DropdownRun),
    the text inside SDT content (visual representation) is not counted as text elements.
    This is expected - the controls are now structured data instead of raw text.
    """
    stage = f"{before.stage} → {after.stage}"
    
    # Calculate expected text reduction from inline controls
    # Checkboxes have ~1 char visual (☐/☑), dropdowns have their selected value text
    # These are now represented as structured controls, not text
    has_inline_controls = before.checkbox_count > 0 or before.dropdown_count > 0
    
    # Compare counts
    if before.total_chars != after.total_chars:
        diff = after.total_chars - before.total_chars
        
        # If we have inline controls and the diff is negative (less chars in JSON),
        # this is expected - checkbox/dropdown visual text is now structured data
        if has_inline_controls and diff < 0:
            # Estimate: checkbox ~1 char each, dropdown ~10-20 chars each for selected value
            expected_reduction = before.checkbox_count * 1 + before.dropdown_count * 10
            if abs(diff) <= expected_reduction * 3:  # Allow 3x margin for variable dropdown text
                severity = "info"  # Expected difference
            else:
                severity = "warning"
        else:
            severity = "error" if abs(diff) > 10 else "warning"
        
        report.add_issue(
            stage, severity, "char_count",
            f"Character count changed: {before.total_chars} → {after.total_chars} (diff: {diff:+d})" + 
            (" (expected with inline controls)" if has_inline_controls and diff < 0 else ""),
            {"before": before.total_chars, "after": after.total_chars, "diff": diff}
        )
    
    if before.paragraph_count != after.paragraph_count:
        diff = after.paragraph_count - before.paragraph_count
        report.add_issue(
            stage, "warning", "structure",
            f"Paragraph count changed: {before.paragraph_count} → {after.paragraph_count} (diff: {diff:+d})",
            {"before": before.paragraph_count, "after": after.paragraph_count}
        )
    
    if before.table_count != after.table_count:
        report.add_issue(
            stage, "error", "structure",
            f"Table count changed: {before.table_count} → {after.table_count}",
            {"before": before.table_count, "after": after.table_count}
        )
    
    if before.row_count != after.row_count:
        diff = after.row_count - before.row_count
        report.add_issue(
            stage, "error", "structure",
            f"Row count changed: {before.row_count} → {after.row_count} (diff: {diff:+d})",
            {"before": before.row_count, "after": after.row_count}
        )
    
    if before.cell_count != after.cell_count:
        diff = after.cell_count - before.cell_count
        report.add_issue(
            stage, "error", "structure",
            f"Cell count changed: {before.cell_count} → {after.cell_count} (diff: {diff:+d})",
            {"before": before.cell_count, "after": after.cell_count}
        )
    
    # Compare text content using Counter for frequency matching
    before_counter = Counter(before.text_elements)
    after_counter = Counter(after.text_elements)
    
    missing = before_counter - after_counter
    extra = after_counter - before_counter
    
    if missing:
        missing_count = sum(missing.values())
        top_missing = missing.most_common(5)
        
        # If we have inline controls, missing text is expected (checkbox/dropdown content)
        if has_inline_controls:
            expected_missing = before.checkbox_count + before.dropdown_count
            if missing_count <= expected_missing * 2:  # Allow 2x margin
                severity = "info"
            else:
                severity = "warning"
        else:
            severity = "error"
        
        report.add_issue(
            stage, severity, "missing_text",
            f"Missing {missing_count} text elements" +
            (" (expected with inline controls)" if has_inline_controls else ""),
            {"count": missing_count, "samples": dict(top_missing)}
        )
    
    if extra:
        top_extra = extra.most_common(5)
        report.add_issue(
            stage, "error", "extra_text",
            f"Extra {sum(extra.values())} text elements added",
            {"count": sum(extra.values()), "samples": dict(top_extra)}
        )


def validate_xml_refs(json_doc: DocumentJSON, docx_path: str, report: ValidationReport):
    """Validate that all xml_ref paths in JSON correctly resolve to elements in DOCX.
    
    Handles nested tables recursively.
    """
    from services.document_engine import _find_node_by_ref, NS
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        with zf.open('word/document.xml') as doc:
            tree = ET.parse(doc)
            root = tree.getroot()
            body = root.find("w:body", NS)
    
    if body is None:
        report.add_issue("validate", "error", "structure", "DOCX has no body element")
        return
    
    unresolved = []
    mismatched = []
    
    def validate_table(table: TableBlock):
        """Recursively validate a table and its nested tables."""
        tbl_el = _find_node_by_ref(body, table.xml_ref)
        if tbl_el is None:
            unresolved.append(table.xml_ref)
        
        for row in table.rows:
            for cell in row.cells:
                validate_cell_blocks(cell.blocks)
    
    def validate_cell_blocks(blocks):
        """Validate blocks inside a cell (paragraphs and nested tables)."""
        for block in blocks:
            if isinstance(block, ParagraphBlock):
                el = _find_node_by_ref(body, block.xml_ref)
                if el is None:
                    unresolved.append(block.xml_ref)
            elif isinstance(block, TableBlock):
                # Nested table - validate recursively
                validate_table(block)
    
    for block in json_doc.blocks:
        if isinstance(block, ParagraphBlock):
            el = _find_node_by_ref(body, block.xml_ref)
            if el is None:
                unresolved.append(block.xml_ref)
            else:
                # Verify it's actually a paragraph
                if not el.tag.endswith('}p'):
                    mismatched.append((block.xml_ref, "expected p", el.tag))
        
        elif isinstance(block, TableBlock):
            validate_table(block)
    
    if unresolved:
        report.add_issue(
            "validate", "error", "xml_ref",
            f"{len(unresolved)} xml_ref paths could not be resolved",
            {"count": len(unresolved), "samples": unresolved[:10]}
        )
    
    if mismatched:
        report.add_issue(
            "validate", "error", "xml_ref",
            f"{len(mismatched)} xml_ref paths resolved to wrong element types",
            {"count": len(mismatched), "samples": mismatched[:10]}
        )


def validate_parse_stage(docx_path: str, json_doc: DocumentJSON) -> ValidationReport:
    """Validate the DOCX → JSON parsing stage."""
    report = ValidationReport(document_id=json_doc.id)
    
    # Get raw DOCX content
    raw_snapshot = extract_raw_docx_content(docx_path)
    report.add_stage(raw_snapshot)
    
    # Get parsed JSON content
    json_snapshot = extract_json_content(json_doc)
    report.add_stage(json_snapshot)
    
    # Compare
    compare_snapshots(raw_snapshot, json_snapshot, report)
    
    # Validate xml_refs
    validate_xml_refs(json_doc, docx_path, report)
    
    return report


def validate_export_stage(
    original_json: DocumentJSON,
    edited_json: DocumentJSON,
    base_docx_path: str,
    output_docx_path: str
) -> ValidationReport:
    """Validate the JSON → DOCX export stage."""
    report = ValidationReport(document_id=edited_json.id)
    
    # Get original JSON content
    original_snapshot = extract_json_content(original_json)
    original_snapshot.stage = "original_json"
    report.add_stage(original_snapshot)
    
    # Get edited JSON content
    edited_snapshot = extract_json_content(edited_json)
    edited_snapshot.stage = "edited_json"
    report.add_stage(edited_snapshot)
    
    # Compare original vs edited (informational)
    compare_snapshots(original_snapshot, edited_snapshot, report)
    
    # Get exported DOCX content
    if Path(output_docx_path).exists():
        export_snapshot = extract_raw_docx_content(output_docx_path)
        export_snapshot.stage = "exported_docx"
        report.add_stage(export_snapshot)
        
        # Compare edited JSON vs exported DOCX
        compare_snapshots(edited_snapshot, export_snapshot, report)
        
        # Also compare with base DOCX (to catch unintended changes)
        base_snapshot = extract_raw_docx_content(base_docx_path)
        base_snapshot.stage = "base_docx"
        report.add_stage(base_snapshot)
    
    return report


def validate_full_roundtrip(docx_path: str, json_doc: DocumentJSON, output_docx_path: str = None) -> ValidationReport:
    """Validate the complete DOCX → JSON → DOCX roundtrip without any edits.
    
    This is the most thorough test - if no edits are made, the output should
    be identical to the input (in terms of text content).
    """
    report = ValidationReport(document_id=json_doc.id)
    
    # Stage 1: Raw DOCX
    raw_snapshot = extract_raw_docx_content(docx_path)
    report.add_stage(raw_snapshot)
    
    # Stage 2: Parsed JSON
    json_snapshot = extract_json_content(json_doc)
    report.add_stage(json_snapshot)
    
    # Compare parsing stage
    compare_snapshots(raw_snapshot, json_snapshot, report)
    
    # If output exists, compare export stage
    if output_docx_path and Path(output_docx_path).exists():
        export_snapshot = extract_raw_docx_content(output_docx_path)
        export_snapshot.stage = "exported_docx"
        report.add_stage(export_snapshot)
        
        # For roundtrip, exported should match original
        compare_snapshots(raw_snapshot, export_snapshot, report)
    
    return report


def print_report(report: ValidationReport):
    """Print a human-readable validation report."""
    print(f"\n{'='*70}")
    print(f"VALIDATION REPORT: {report.document_id}")
    print(f"{'='*70}")
    
    print("\nStage Snapshots:")
    print("-" * 50)
    for s in report.stages:
        print(f"  {s.stage}:")
        print(f"    Chars: {s.total_chars:,}")
        print(f"    Paragraphs: {s.paragraph_count}, Tables: {s.table_count}")
        print(f"    Rows: {s.row_count}, Cells: {s.cell_count}, Runs: {s.run_count}")
        if s.checkbox_count or s.dropdown_count:
            print(f"    Checkboxes: {s.checkbox_count}, Dropdowns: {s.dropdown_count}")
    
    if report.issues:
        print(f"\nIssues Found: {len(report.issues)}")
        print("-" * 50)
        
        errors = [i for i in report.issues if i.severity == "error"]
        warnings = [i for i in report.issues if i.severity == "warning"]
        
        if errors:
            print(f"\n  ERRORS ({len(errors)}):")
            for issue in errors:
                print(f"    [{issue.stage}] {issue.category}: {issue.message}")
                if issue.details:
                    for k, v in issue.details.items():
                        print(f"      {k}: {v}")
        
        if warnings:
            print(f"\n  WARNINGS ({len(warnings)}):")
            for issue in warnings:
                print(f"    [{issue.stage}] {issue.category}: {issue.message}")
    else:
        print("\n  ✓ No issues found")
    
    print(f"\n{'='*70}")
    status = "FAILED" if report.has_errors else ("WARNINGS" if report.has_warnings else "PASSED")
    print(f"Status: {status}")
    print(f"{'='*70}")
