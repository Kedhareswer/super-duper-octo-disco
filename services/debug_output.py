"""Debug Output Service - Saves intermediate results for document processing.

This module saves intermediate results during document processing for debugging
and verification purposes. Each document gets its own folder with:
- Original document.xml (extracted from DOCX)
- Parsed JSON structure
- Validation reports
- Export comparisons
"""

from __future__ import annotations

import json
import zipfile
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any, List
from dataclasses import dataclass, asdict
from xml.etree import ElementTree as ET
from xml.dom import minidom

from models.schemas import DocumentJSON

logger = logging.getLogger(__name__)


# Debug output root directory
DEBUG_ROOT = Path("data/debug")


@dataclass
class DebugSnapshot:
    """Snapshot of a processing stage for debugging."""
    stage: str
    timestamp: str
    data_file: str
    summary: Dict[str, Any]
    issues: List[str]


def _ensure_debug_dir(document_id: str) -> Path:
    """Create and return the debug directory for a document."""
    # Sanitize document_id for filesystem
    safe_id = document_id.replace("/", "_").replace("\\", "_").replace(":", "_")
    debug_dir = DEBUG_ROOT / safe_id
    debug_dir.mkdir(parents=True, exist_ok=True)
    return debug_dir


def _prettify_xml(xml_bytes: bytes) -> str:
    """Pretty-print XML for readability."""
    try:
        # Parse and re-serialize with indentation
        root = ET.fromstring(xml_bytes)
        rough_string = ET.tostring(root, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")
    except Exception:
        # If prettify fails, return as-is
        return xml_bytes.decode('utf-8', errors='replace')


def save_docx_structure(document_id: str, docx_path: str) -> Dict[str, Any]:
    """Extract and save the DOCX ZIP structure and key XML files.
    
    Returns a summary of the DOCX structure.
    """
    debug_dir = _ensure_debug_dir(document_id)
    summary = {
        "docx_path": docx_path,
        "files": [],
        "document_xml_size": 0,
        "total_files": 0,
    }
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            summary["total_files"] = len(zf.namelist())
            
            # Save list of all files
            files_info = []
            for name in sorted(zf.namelist()):
                info = zf.getinfo(name)
                files_info.append({
                    "name": name,
                    "size": info.file_size,
                    "compressed_size": info.compress_size,
                })
            
            summary["files"] = files_info
            
            # Save structure manifest
            manifest_path = debug_dir / "01_docx_structure.json"
            manifest_path.write_text(json.dumps({
                "document_id": document_id,
                "docx_path": docx_path,
                "files": files_info,
            }, indent=2))
            
            # Extract and save key XML files
            key_files = [
                ('word/document.xml', '02_document.xml'),
                ('[Content_Types].xml', '03_content_types.xml'),
                ('word/styles.xml', '04_styles.xml'),
                ('word/_rels/document.xml.rels', '05_document_rels.xml'),
            ]
            
            for src_name, dest_name in key_files:
                if src_name in zf.namelist():
                    with zf.open(src_name) as f:
                        content = f.read()
                        
                        if src_name == 'word/document.xml':
                            summary["document_xml_size"] = len(content)
                        
                        # Save raw XML
                        raw_path = debug_dir / dest_name
                        raw_path.write_bytes(content)
                        
                        # Save prettified version
                        pretty_path = debug_dir / dest_name.replace('.xml', '_pretty.xml')
                        try:
                            pretty_content = _prettify_xml(content)
                            pretty_path.write_text(pretty_content, encoding='utf-8')
                        except Exception as e:
                            logger.warning(f"Could not prettify {src_name}: {e}")
            
            logger.info(f"[DEBUG] Saved DOCX structure for {document_id}: {summary['total_files']} files")
            
    except Exception as e:
        summary["error"] = str(e)
        logger.error(f"[DEBUG] Failed to extract DOCX structure for {document_id}: {e}")
    
    return summary


def save_parsed_json(document_id: str, json_doc: DocumentJSON) -> Dict[str, Any]:
    """Save the parsed JSON structure with summary statistics.
    
    Returns a summary of the parsed content.
    """
    debug_dir = _ensure_debug_dir(document_id)
    
    # Calculate statistics
    stats = {
        "block_count": len(json_doc.blocks),
        "paragraph_count": 0,
        "table_count": 0,
        "drawing_count": 0,
        "run_count": 0,
        "total_chars": 0,
        "checkbox_count": len(json_doc.checkboxes),
        "dropdown_count": len(json_doc.dropdowns),
        "cell_count": 0,
        "nested_table_count": 0,
    }
    
    from models.schemas import ParagraphBlock, TableBlock, DrawingBlock
    
    def count_in_blocks(blocks):
        """Recursively count elements in blocks."""
        for block in blocks:
            if isinstance(block, ParagraphBlock):
                stats["paragraph_count"] += 1
                for run in block.runs:
                    stats["run_count"] += 1
                    # Only TextRuns have text attribute; CheckboxRun/DropdownRun don't
                    if hasattr(run, 'text') and run.text:
                        stats["total_chars"] += len(run.text)
            elif isinstance(block, TableBlock):
                stats["table_count"] += 1
                for row in block.rows:
                    for cell in row.cells:
                        stats["cell_count"] += 1
                        count_in_blocks(cell.blocks)
                        # Check for nested tables
                        for cb in cell.blocks:
                            if isinstance(cb, TableBlock):
                                stats["nested_table_count"] += 1
            elif isinstance(block, DrawingBlock):
                stats["drawing_count"] += 1
    
    count_in_blocks(json_doc.blocks)
    
    # Save full JSON (indented for readability)
    json_path = debug_dir / "10_parsed_json.json"
    json_path.write_text(json_doc.model_dump_json(indent=2), encoding='utf-8')
    
    # Save compact JSON (for size comparison)
    compact_path = debug_dir / "11_parsed_json_compact.json"
    compact_path.write_text(json_doc.model_dump_json(), encoding='utf-8')
    
    # Save statistics
    stats_path = debug_dir / "12_parse_stats.json"
    stats_path.write_text(json.dumps({
        "document_id": document_id,
        "stats": stats,
        "checkboxes": [{"id": cb.id, "label": cb.label, "checked": cb.checked} for cb in json_doc.checkboxes],
        "dropdowns": [{"id": dd.id, "label": dd.label, "options": dd.options, "selected": dd.selected} for dd in json_doc.dropdowns],
    }, indent=2), encoding='utf-8')
    
    logger.info(f"[DEBUG] Saved parsed JSON for {document_id}: {stats['total_chars']} chars, {stats['paragraph_count']} paragraphs, {stats['table_count']} tables")
    
    return stats


def save_validation_report(document_id: str, stage: str, report: Dict[str, Any]) -> None:
    """Save a validation report for a processing stage."""
    debug_dir = _ensure_debug_dir(document_id)
    
    # Determine filename based on stage
    stage_files = {
        "parse": "20_validation_parse.json",
        "export": "21_validation_export.json",
        "roundtrip": "22_validation_roundtrip.json",
    }
    
    filename = stage_files.get(stage, f"20_validation_{stage}.json")
    report_path = debug_dir / filename
    report_path.write_text(json.dumps(report, indent=2, default=str), encoding='utf-8')
    
    logger.info(f"[DEBUG] Saved validation report ({stage}) for {document_id}")


def save_export_comparison(
    document_id: str, 
    original_docx_path: str, 
    exported_docx_path: str
) -> Dict[str, Any]:
    """Compare original and exported DOCX files and save the comparison.
    
    Returns a summary of differences.
    """
    debug_dir = _ensure_debug_dir(document_id)
    comparison = {
        "original_path": original_docx_path,
        "exported_path": exported_docx_path,
        "file_differences": [],
        "document_xml_diff": None,
    }
    
    try:
        with zipfile.ZipFile(original_docx_path, 'r') as orig_zf:
            with zipfile.ZipFile(exported_docx_path, 'r') as exp_zf:
                orig_files = set(orig_zf.namelist())
                exp_files = set(exp_zf.namelist())
                
                # Check for missing/added files
                missing = orig_files - exp_files
                added = exp_files - orig_files
                
                if missing:
                    comparison["file_differences"].append({
                        "type": "missing_in_export",
                        "files": list(missing),
                    })
                
                if added:
                    comparison["file_differences"].append({
                        "type": "added_in_export",
                        "files": list(added),
                    })
                
                # Compare document.xml
                with orig_zf.open('word/document.xml') as f:
                    orig_doc = f.read()
                with exp_zf.open('word/document.xml') as f:
                    exp_doc = f.read()
                
                comparison["document_xml_diff"] = {
                    "original_size": len(orig_doc),
                    "exported_size": len(exp_doc),
                    "size_difference": len(exp_doc) - len(orig_doc),
                    "identical": orig_doc == exp_doc,
                }
                
                # Save exported document.xml for comparison
                exp_doc_path = debug_dir / "30_exported_document.xml"
                exp_doc_path.write_bytes(exp_doc)
                
                try:
                    pretty_exp = _prettify_xml(exp_doc)
                    pretty_path = debug_dir / "31_exported_document_pretty.xml"
                    pretty_path.write_text(pretty_exp, encoding='utf-8')
                except Exception:
                    pass
        
        # Save comparison report
        comp_path = debug_dir / "32_export_comparison.json"
        comp_path.write_text(json.dumps(comparison, indent=2), encoding='utf-8')
        
        logger.info(f"[DEBUG] Saved export comparison for {document_id}")
        
    except Exception as e:
        comparison["error"] = str(e)
        logger.error(f"[DEBUG] Failed to compare exports for {document_id}: {e}")
    
    return comparison


def save_edit_snapshot(
    document_id: str, 
    edit_id: str,
    before_json: DocumentJSON,
    after_json: DocumentJSON,
    edit_details: Dict[str, Any]
) -> None:
    """Save before/after snapshots of an edit operation."""
    debug_dir = _ensure_debug_dir(document_id)
    edits_dir = debug_dir / "edits"
    edits_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    edit_dir = edits_dir / f"{timestamp}_{edit_id}"
    edit_dir.mkdir(exist_ok=True)
    
    # Save before JSON
    before_path = edit_dir / "before.json"
    before_path.write_text(before_json.model_dump_json(indent=2), encoding='utf-8')
    
    # Save after JSON
    after_path = edit_dir / "after.json"
    after_path.write_text(after_json.model_dump_json(indent=2), encoding='utf-8')
    
    # Save edit details
    details_path = edit_dir / "edit_details.json"
    details_path.write_text(json.dumps(edit_details, indent=2, default=str), encoding='utf-8')
    
    logger.info(f"[DEBUG] Saved edit snapshot for {document_id}/{edit_id}")


def create_debug_manifest(document_id: str) -> Dict[str, Any]:
    """Create a manifest of all debug files for a document."""
    debug_dir = _ensure_debug_dir(document_id)
    
    manifest = {
        "document_id": document_id,
        "debug_dir": str(debug_dir),
        "created_at": datetime.now().isoformat(),
        "files": [],
    }
    
    for file_path in sorted(debug_dir.glob("*")):
        if file_path.is_file():
            manifest["files"].append({
                "name": file_path.name,
                "size": file_path.stat().st_size,
            })
    
    # Save manifest
    manifest_path = debug_dir / "00_manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding='utf-8')
    
    return manifest


def get_debug_dir(document_id: str) -> Path:
    """Get the debug directory path for a document."""
    return _ensure_debug_dir(document_id)
