"""XLSX Writer - Applies JSON edits back to Excel files.

Preserves structural fidelity by:
1. Copying base XLSX structure byte-for-byte
2. Only modifying cell values and specific editable content
3. Maintaining all formatting, styles, images, and complex elements
4. Preserving ALL namespace declarations (critical for Excel compatibility)
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET

from .schemas import (
    CellDataType,
    ExcelCellJSON,
    ExcelSheetJSON,
    ExcelWorkbookJSON,
    SharedStringItem,
)
from .parser import NS, col_index_to_letter, parse_cell_ref


# Register ALL known Excel namespaces for output
# This is critical - ElementTree drops unregistered namespaces
_ALL_EXCEL_NAMESPACES = {
    "": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision10",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "vml": "urn:schemas-microsoft-com:vml",
    "x": "urn:schemas-microsoft-com:office:excel",
    "o": "urn:schemas-microsoft-com:office:office",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

for prefix, uri in _ALL_EXCEL_NAMESPACES.items():
    ET.register_namespace(prefix, uri)

# Also register from parser NS
for prefix, uri in NS.items():
    ET.register_namespace(prefix if prefix != "main" else "", uri)


def _extract_root_tag(xml_bytes: bytes) -> Tuple[bytes, bytes, bytes]:
    """Extract the original root element opening/closing tags from XML.
    
    Returns:
        (xml_declaration, root_open_tag, root_close_tag)
        
    This preserves ALL original namespace declarations which is critical
    for Excel compatibility. ElementTree often drops unused namespace
    declarations, but Excel requires them (e.g., for mc:Ignorable).
    """
    xml_str = xml_bytes.decode('utf-8')
    
    # Find XML declaration
    decl_match = re.match(r'(<\?xml[^?]*\?>)\s*', xml_str)
    if decl_match:
        xml_decl = decl_match.group(1).encode('utf-8')
        rest = xml_str[decl_match.end():]
    else:
        xml_decl = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        rest = xml_str
    
    # Find opening root tag (worksheet, sst, etc.)
    # Match from < to first > that's not part of an attribute
    root_match = re.match(r'(<[a-zA-Z][^>]*>)', rest)
    if root_match:
        root_open = root_match.group(1).encode('utf-8')
    else:
        root_open = b'<worksheet>'
    
    # Find closing tag
    close_match = re.search(r'(</[a-zA-Z][^>]*>)\s*$', xml_str)
    if close_match:
        root_close = close_match.group(1).encode('utf-8')
    else:
        root_close = b'</worksheet>'
    
    return xml_decl, root_open, root_close


def _serialize_element_inner(element: ET.Element, ns: str) -> bytes:
    """Serialize an element's inner content (children only, no root tag).
    
    This allows us to preserve the original root tag with all namespaces
    while still using ElementTree for content manipulation.
    
    Note: We strip redundant default namespace declarations that ElementTree
    adds to child elements - the namespace is already declared on the root.
    """
    buffer = BytesIO()
    
    # Write each child element
    for child in element:
        child_str = ET.tostring(child, encoding='unicode')
        # Remove redundant default namespace declarations added by ElementTree
        # These are already declared on the root element
        child_str = child_str.replace(f' xmlns="{ns}"', '')
        child_str = child_str.replace(f" xmlns='{ns}'", '')
        buffer.write(child_str.encode('utf-8'))
    
    return buffer.getvalue()


def _get_shared_string_index(
    value: str,
    shared_strings: List[SharedStringItem],
    new_strings: Dict[str, int],
) -> int:
    """Get or create a shared string index for a value."""
    # Check existing
    for ss in shared_strings:
        if ss.text == value:
            return ss.index
    
    # Check newly added
    if value in new_strings:
        return new_strings[value]
    
    # Add new
    new_index = len(shared_strings) + len(new_strings)
    new_strings[value] = new_index
    return new_index


def _update_shared_strings(
    zf_in: zipfile.ZipFile,
    new_strings: Dict[str, int],
) -> bytes:
    """Update shared strings XML with new entries.
    
    Preserves original root element structure for Excel compatibility.
    """
    original_xml: Optional[bytes] = None
    
    try:
        original_xml = zf_in.read("xl/sharedStrings.xml")
        tree = ET.parse(BytesIO(original_xml))
        root = tree.getroot()
    except KeyError:
        # Create new shared strings file
        root = ET.Element(f"{{{NS['main']}}}sst")
        root.set("xmlns", NS["main"])
        tree = ET.ElementTree(root)
    
    ns = NS["main"]
    
    # Get current count
    current_count = int(root.get("count", 0))
    unique_count = int(root.get("uniqueCount", 0))
    
    # Add new strings
    for text, index in sorted(new_strings.items(), key=lambda x: x[1]):
        si = ET.SubElement(root, f"{{{ns}}}si")
        t = ET.SubElement(si, f"{{{ns}}}t")
        t.text = text
        # Preserve whitespace
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    
    # Update counts in root element
    new_count = current_count + len(new_strings)
    new_unique = unique_count + len(new_strings)
    root.set("count", str(new_count))
    root.set("uniqueCount", str(new_unique))
    
    # Serialize - if we have original XML, try to preserve its structure
    if original_xml:
        # Extract original root tag and modify counts
        xml_decl, root_open, root_close = _extract_root_tag(original_xml)
        
        # Update counts in the root_open tag
        root_open_str = root_open.decode('utf-8')
        root_open_str = re.sub(r'count="[^"]*"', f'count="{new_count}"', root_open_str)
        root_open_str = re.sub(r'uniqueCount="[^"]*"', f'uniqueCount="{new_unique}"', root_open_str)
        root_open = root_open_str.encode('utf-8')
        
        # Serialize inner content
        inner_content = _serialize_element_inner(root, ns)
        
        result = xml_decl + b"\r\n" + root_open + inner_content + root_close
    else:
        # New file - just serialize normally
        buffer = BytesIO()
        tree.write(buffer, xml_declaration=True, encoding="UTF-8")
        result = buffer.getvalue()
        
        # Fix XML declaration for Excel compatibility
        result = result.replace(
            b"<?xml version='1.0' encoding='UTF-8'?>",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        )
    
    return result


def _apply_cell_changes(
    sheet_xml: bytes,
    sheet: ExcelSheetJSON,
    original_shared_strings: List[SharedStringItem],
    new_strings: Dict[str, int],
) -> bytes:
    """Apply cell value changes to a worksheet XML.
    
    Only updates cells that have been marked as dirty.
    """
    
    tree = ET.parse(BytesIO(sheet_xml))
    root = tree.getroot()
    ns = NS["main"]
    
    # Build lookup of ONLY dirty cells
    cell_updates: Dict[str, ExcelCellJSON] = {
        cell.ref: cell for cell in sheet.cells
        if cell.dirty
    }
    
    # If no dirty cells, return original XML unchanged
    if not cell_updates:
        return sheet_xml
    
    # Find sheetData
    sheet_data = root.find(f"{{{ns}}}sheetData")
    if sheet_data is None:
        return sheet_xml
    
    # Track which cells we've updated
    updated_refs: Set[str] = set()
    
    # Update existing cells
    for row_el in sheet_data.findall(f"{{{ns}}}row"):
        for cell_el in row_el.findall(f"{{{ns}}}c"):
            cell_ref = cell_el.get("r")
            if cell_ref not in cell_updates:
                continue
            
            cell = cell_updates[cell_ref]
            updated_refs.add(cell_ref)
            
            # Get current type
            current_type = cell_el.get("t")
            
            # Update value
            v_el = cell_el.find(f"{{{ns}}}v")
            
            if cell.value is None:
                # Clear value
                if v_el is not None:
                    cell_el.remove(v_el)
                continue
            
            # Create v element if needed
            if v_el is None:
                v_el = ET.SubElement(cell_el, f"{{{ns}}}v")
            
            # Handle different value types
            if isinstance(cell.value, bool):
                cell_el.set("t", "b")
                v_el.text = "1" if cell.value else "0"
            elif isinstance(cell.value, (int, float)):
                # Remove type attribute for numbers
                if "t" in cell_el.attrib:
                    del cell_el.attrib["t"]
                v_el.text = str(cell.value)
            elif isinstance(cell.value, str):
                # Use shared strings for text
                ss_index = _get_shared_string_index(
                    cell.value,
                    original_shared_strings,
                    new_strings,
                )
                cell_el.set("t", "s")
                v_el.text = str(ss_index)
            else:
                # Convert to string
                ss_index = _get_shared_string_index(
                    str(cell.value),
                    original_shared_strings,
                    new_strings,
                )
                cell_el.set("t", "s")
                v_el.text = str(ss_index)
    
    # Add new dirty cells that don't exist in the original
    for cell in sheet.cells:
        if cell.ref in updated_refs:
            continue
        
        # Only add cells that are dirty and have values
        if not cell.dirty:
            continue
        
        if cell.value is None:
            continue
        
        # Find or create the row
        row_el = None
        for r in sheet_data.findall(f"{{{ns}}}row"):
            if int(r.get("r", 0)) == cell.row:
                row_el = r
                break
        
        if row_el is None:
            # Create new row
            row_el = ET.SubElement(sheet_data, f"{{{ns}}}row")
            row_el.set("r", str(cell.row))
        
        # Create new cell
        cell_el = ET.SubElement(row_el, f"{{{ns}}}c")
        cell_el.set("r", cell.ref)
        
        if cell.style_index is not None:
            cell_el.set("s", str(cell.style_index))
        
        v_el = ET.SubElement(cell_el, f"{{{ns}}}v")
        
        if isinstance(cell.value, bool):
            cell_el.set("t", "b")
            v_el.text = "1" if cell.value else "0"
        elif isinstance(cell.value, (int, float)):
            v_el.text = str(cell.value)
        elif isinstance(cell.value, str):
            ss_index = _get_shared_string_index(
                cell.value,
                original_shared_strings,
                new_strings,
            )
            cell_el.set("t", "s")
            v_el.text = str(ss_index)
    
    # Serialize while preserving original root tag with all namespaces
    # This is critical because ElementTree drops "unused" namespace declarations
    # but Excel requires them (e.g., mc:Ignorable references xr, xr2, xr3, x14ac)
    
    xml_decl, root_open, root_close = _extract_root_tag(sheet_xml)
    
    # Serialize inner content (all children of root)
    inner_content = _serialize_element_inner(root, ns)
    
    # Reassemble with original root tag
    result = xml_decl + b"\r\n" + root_open + inner_content + root_close
    
    return result


def _get_sheet_path_map(zf: zipfile.ZipFile) -> Dict[str, str]:
    """Get mapping from sheet name to sheet file path."""
    
    # Parse workbook.xml for sheet info
    with zf.open("xl/workbook.xml") as f:
        wb_tree = ET.parse(f)
        wb_root = wb_tree.getroot()
    
    ns = NS["main"]
    r_ns = NS["r"]
    
    # Get sheet names and r:ids
    sheets_el = wb_root.find(f"{{{ns}}}sheets")
    sheet_infos: Dict[str, str] = {}  # name -> r_id
    
    if sheets_el is not None:
        for sheet in sheets_el.findall(f"{{{ns}}}sheet"):
            name = sheet.get("name")
            r_id = sheet.get(f"{{{r_ns}}}id")
            if name and r_id:
                sheet_infos[name] = r_id
    
    # Parse workbook rels to get sheet paths
    with zf.open("xl/_rels/workbook.xml.rels") as f:
        rels_tree = ET.parse(f)
        rels_root = rels_tree.getroot()
    
    ns_rel = NS["rel"]
    id_to_target: Dict[str, str] = {}
    
    for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            id_to_target[rel_id] = target
    
    # Build name -> path mapping
    result: Dict[str, str] = {}
    for name, r_id in sheet_infos.items():
        target = id_to_target.get(r_id, "")
        if target.startswith("/"):
            path = target[1:]
        else:
            path = f"xl/{target}"
        result[name] = path
    
    return result


def apply_json_to_xlsx(
    json_doc: ExcelWorkbookJSON,
    base_xlsx_path: str,
    out_xlsx_path: str,
) -> str:
    """Apply JSON edits to an XLSX file and save to a new path.
    
    Strategy:
    1. Copy all files from base XLSX byte-for-byte
    2. Only update cell values that have actually changed
    3. Update shared strings with any new text values
    4. Preserve all other structure (styles, images, etc.)
    
    Returns: Path to the output file
    """
    import shutil
    
    base = Path(base_xlsx_path)
    out = Path(out_xlsx_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    
    # Check if there are any actual edits
    has_edits = any(
        cell.dirty for sheet in json_doc.sheets for cell in sheet.cells
        if cell.dirty
    )
    
    # If no edits, just copy the file directly
    if not has_edits:
        shutil.copy2(base, out)
        return str(out)
    
    # Track new shared strings
    new_strings: Dict[str, int] = {}
    
    # Build sheet lookup - only for sheets with dirty cells
    sheet_by_name: Dict[str, ExcelSheetJSON] = {}
    for s in json_doc.sheets:
        dirty_cells = [c for c in s.cells if c.dirty]
        if dirty_cells:
            sheet_by_name[s.name] = s
    
    # If no dirty sheets, just copy
    if not sheet_by_name:
        shutil.copy2(base, out)
        return str(out)
    
    with zipfile.ZipFile(base, "r") as zf_in:
        # Get sheet path mapping
        sheet_paths = _get_sheet_path_map(zf_in)
        
        # Files to update
        updated_files: Dict[str, bytes] = {}
        
        # Update only sheets with dirty cells
        for sheet_name, sheet_path in sheet_paths.items():
            if sheet_name not in sheet_by_name:
                continue
            
            sheet = sheet_by_name[sheet_name]
            
            try:
                sheet_xml = zf_in.read(sheet_path)
                updated_xml = _apply_cell_changes(
                    sheet_xml,
                    sheet,
                    json_doc.shared_strings,
                    new_strings,
                )
                updated_files[sheet_path] = updated_xml
            except KeyError:
                continue
        
        # Update shared strings if we added new ones
        if new_strings:
            updated_ss = _update_shared_strings(zf_in, new_strings)
            updated_files["xl/sharedStrings.xml"] = updated_ss
        
        # Write output file - copy base and only overwrite changed files
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename in updated_files:
                    # Write updated content with same compression
                    zf_out.writestr(
                        item.filename, 
                        updated_files[item.filename],
                        compress_type=item.compress_type or zipfile.ZIP_DEFLATED
                    )
                else:
                    # Copy original bytes exactly
                    data = zf_in.read(item.filename)
                    zf_out.writestr(
                        item.filename, 
                        data,
                        compress_type=item.compress_type or zipfile.ZIP_DEFLATED
                    )
    
    return str(out)


def apply_cell_edits(
    json_doc: ExcelWorkbookJSON,
    edits: List[Dict[str, Any]],
) -> ExcelWorkbookJSON:
    """Apply a list of cell edits to the workbook JSON.
    
    Each edit should have:
    - sheet: sheet name or index
    - cell: cell reference (e.g., "A1")
    - value: new value
    
    Returns: Updated workbook JSON (modifies in place)
    """
    
    for edit in edits:
        sheet_ref = edit.get("sheet")
        cell_ref = edit.get("cell")
        new_value = edit.get("value")
        
        if sheet_ref is None or cell_ref is None:
            continue
        
        # Find sheet
        sheet: Optional[ExcelSheetJSON] = None
        if isinstance(sheet_ref, int):
            sheet = json_doc.get_sheet_by_index(sheet_ref)
        else:
            sheet = json_doc.get_sheet(str(sheet_ref))
        
        if sheet is None:
            continue
        
        # Find or create cell
        existing_cell = sheet.get_cell(cell_ref)
        
        if existing_cell:
            existing_cell.value = new_value
        else:
            # Create new cell
            col_letter, col, row = parse_cell_ref(cell_ref)
            new_cell = ExcelCellJSON(
                id=f"{sheet.id}-{cell_ref}",
                ref=cell_ref,
                row=row,
                col=col,
                col_letter=col_letter,
                value=new_value,
            )
            sheet.cells.append(new_cell)
    
    return json_doc
