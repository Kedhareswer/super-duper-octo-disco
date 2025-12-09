"""Document Engine - DOCX to JSON conversion and back.

This module handles:
1. Parsing DOCX files into an editable JSON structure (DocumentJSON)
2. Applying JSON edits back to DOCX files
3. Preserving formatting, form fields, and document structure
"""
from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import List
# Use lxml for proper namespace prefix preservation in DOCX export
# ElementTree corrupts namespace prefixes causing "Word found unreadable content" errors
try:
    from lxml import etree as ET
    USING_LXML = True
except ImportError:
    from xml.etree import ElementTree as ET
    USING_LXML = False

from models.schemas import (
    CellBorder,
    CellBorders,
    CheckboxField,
    CheckboxRun,
    DocumentJSON,
    DrawingBlock,
    DropdownField,
    DropdownRun,
    InlineContent,
    ParagraphBlock,
    TextRun,
    Run,  # Alias for TextRun
    TableBlock,
    TableCell,
    TableRow,
    ValidationErrorDetail,
    ValidationResult,
)


# Namespaces for OOXML - comprehensive list to preserve prefixes
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "oel": "http://schemas.microsoft.com/office/2019/extlst",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
}

# Register all namespaces with ElementTree to preserve prefixes during serialization
# Note: lxml handles this automatically, but we still register for ElementTree fallback
if not USING_LXML:
    for _prefix, _uri in NS.items():
        ET.register_namespace(_prefix, _uri)


# =============================================================================
# PARSING: DOCX → JSON
# =============================================================================

def _extract_checkboxes(root: ET.Element) -> List[CheckboxField]:
    """Extract all checkbox content controls from the document.
    
    OOXML checkboxes are SDT elements with w14:checkbox inside sdtPr.
    Structure:
    <w:sdt>
      <w:sdtPr>
        <w:alias w:val="field_name"/>
        <w:tag w:val="field_tag"/>
        <w:id w:val="123"/>
        <w14:checkbox>
          <w14:checked w14:val="0|1"/>
          ...
        </w14:checkbox>
      </w:sdtPr>
      <w:sdtContent>...</w:sdtContent>
    </w:sdt>
    """
    checkboxes: List[CheckboxField] = []
    
    # Find all SDT elements
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is None:
            continue
        
        # Check if this is a checkbox
        checkbox_el = sdt_pr.find("w14:checkbox", NS)
        if checkbox_el is None:
            continue
        
        # Extract checkbox properties
        alias_el = sdt_pr.find("w:alias", NS)
        tag_el = sdt_pr.find("w:tag", NS)
        id_el = sdt_pr.find("w:id", NS)
        checked_el = checkbox_el.find("w14:checked", NS)
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "") if id_el is not None else ""
        label = alias_el.attrib.get(f"{{{NS['w']}}}val", "") if alias_el is not None else ""
        tag = tag_el.attrib.get(f"{{{NS['w']}}}val", "") if tag_el is not None else ""
        
        # w14:checked w14:val="0" means unchecked, "1" means checked
        checked = False
        if checked_el is not None:
            checked_val = checked_el.attrib.get(f"{{{NS['w14']}}}val", "0")
            checked = checked_val == "1"
        
        checkboxes.append(CheckboxField(
            id=f"checkbox-{field_id}",
            xml_ref=f"sdt[id={field_id}]",
            label=label or tag or f"Checkbox {field_id}",
            checked=checked,
        ))
    
    return checkboxes


def _extract_dropdowns(root: ET.Element) -> List[DropdownField]:
    """Extract all dropdown/comboBox content controls from the document.
    
    OOXML dropdowns are SDT elements with w:comboBox or w:dropDownList inside sdtPr.
    Structure:
    <w:sdt>
      <w:sdtPr>
        <w:alias w:val="field_name"/>
        <w:tag w:val="field_tag"/>
        <w:id w:val="123"/>
        <w:comboBox>
          <w:listItem w:displayText="Option1" w:value="val1"/>
          <w:listItem w:displayText="Option2" w:value="val2"/>
        </w:comboBox>
      </w:sdtPr>
      <w:sdtContent>
        <w:tc>...<w:t>SelectedValue</w:t>...</w:tc>
      </w:sdtContent>
    </w:sdt>
    """
    dropdowns: List[DropdownField] = []
    
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is None:
            continue
        
        # Check for comboBox or dropDownList
        combo_el = sdt_pr.find("w:comboBox", NS)
        dropdown_el = sdt_pr.find("w:dropDownList", NS)
        list_el = combo_el if combo_el is not None else dropdown_el
        
        if list_el is None:
            continue
        
        # Extract properties
        alias_el = sdt_pr.find("w:alias", NS)
        tag_el = sdt_pr.find("w:tag", NS)
        id_el = sdt_pr.find("w:id", NS)
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "") if id_el is not None else ""
        label = alias_el.attrib.get(f"{{{NS['w']}}}val", "") if alias_el is not None else ""
        tag = tag_el.attrib.get(f"{{{NS['w']}}}val", "") if tag_el is not None else ""
        
        # Extract options
        options: List[str] = []
        for item in list_el.findall("w:listItem", NS):
            display_text = item.attrib.get(f"{{{NS['w']}}}displayText", "")
            value = item.attrib.get(f"{{{NS['w']}}}value", display_text)
            options.append(display_text or value)
        
        # Get currently selected value from sdtContent
        sdt_content = sdt.find("w:sdtContent", NS)
        selected = None
        if sdt_content is not None:
            # Find the text content - preserve original whitespace
            for t_el in sdt_content.iter(f"{{{NS['w']}}}t"):
                if t_el.text:
                    selected = t_el.text  # Don't strip - preserve original
                    break
        
        dropdowns.append(DropdownField(
            id=f"dropdown-{field_id}",
            xml_ref=f"sdt[id={field_id}]",
            label=label or tag or f"Dropdown {field_id}",
            options=options,
            selected=selected,
        ))
    
    return dropdowns


def _extract_drawing(drawing_el: ET.Element, drawing_index: int) -> DrawingBlock | None:
    """Extract drawing information from a w:drawing element."""
    
    # Check for inline or anchor
    inline = drawing_el.find("wp:inline", NS)
    anchor = drawing_el.find("wp:anchor", NS)
    container = inline if inline is not None else anchor
    
    if container is None:
        return None
    
    # Get extent (size) - EMUs to inches (914400 EMUs per inch)
    width_inches = 0.0
    height_inches = 0.0
    extent = container.find("wp:extent", NS)
    if extent is not None:
        cx = int(extent.attrib.get("cx", 0))
        cy = int(extent.attrib.get("cy", 0))
        width_inches = cx / 914400
        height_inches = cy / 914400
    
    # Get name from docPr
    name = None
    doc_pr = container.find("wp:docPr", NS)
    if doc_pr is not None:
        name = doc_pr.attrib.get("name")
    
    # Determine drawing type
    drawing_type = "unknown"
    if container.find(".//wpg:wgp", NS) is not None:
        drawing_type = "vector_group"
    elif container.find(".//pic:pic", NS) is not None:
        drawing_type = "image"
    
    return DrawingBlock(
        id=f"drawing-{drawing_index}",
        xml_ref=f"drawing[{drawing_index}]",
        name=name,
        width_inches=width_inches,
        height_inches=height_inches,
        drawing_type=drawing_type,
    )


def _iter_body_elements(body_el):
    """Yield top-level body elements (paragraphs, tables, and drawings) in document order.
    
    Paragraphs are always yielded to capture text content.
    Drawings are yielded separately after their containing paragraph.
    SDT (content control) elements at body level are unwrapped to process their contents.
    """

    for el in body_el:
        if el.tag == f"{{{NS['w']}}}p":
            # Always yield the paragraph first to capture any text runs
            yield "p", el
            # Then yield any drawings found in the paragraph as separate blocks
            for drawing in el.findall(".//w:drawing", NS):
                yield "drawing", drawing
        elif el.tag == f"{{{NS['w']}}}tbl":
            yield "tbl", el
        elif el.tag == f"{{{NS['w']}}}sdt":
            # SDT (Structured Document Tag / content control) at body level
            # Unwrap and process contents from sdtContent
            sdt_content = el.find("w:sdtContent", NS)
            if sdt_content is not None:
                # Recursively process elements inside SDT
                yield from _iter_body_elements(sdt_content)


def _get_text_from_run(run_el: ET.Element) -> str:
    texts = [t.text or "" for t in run_el.findall("w:t", NS)]
    return "".join(texts)


def _is_bold(run_el: ET.Element) -> bool:
    rpr = run_el.find("w:rPr", NS)
    if rpr is None:
        return False
    return rpr.find("w:b", NS) is not None


def _is_italic(run_el: ET.Element) -> bool:
    rpr = run_el.find("w:rPr", NS)
    if rpr is None:
        return False
    return rpr.find("w:i", NS) is not None


def _get_text_color(run_el: ET.Element) -> str | None:
    """Extract text color from run properties."""
    rpr = run_el.find("w:rPr", NS)
    if rpr is None:
        return None
    color_el = rpr.find("w:color", NS)
    if color_el is None:
        return None
    return color_el.attrib.get(f"{{{NS['w']}}}val")


def _get_cell_background(tc_el: ET.Element) -> str | None:
    """Extract cell background color from cell properties."""
    tc_pr = tc_el.find("w:tcPr", NS)
    if tc_pr is None:
        return None
    shd_el = tc_pr.find("w:shd", NS)
    if shd_el is None:
        return None
    fill = shd_el.attrib.get(f"{{{NS['w']}}}fill")
    # "auto" means no fill
    if fill and fill.lower() != "auto":
        return fill
    return None


def _parse_border(border_el: ET.Element | None) -> CellBorder | None:
    """Parse a single border element."""
    if border_el is None:
        return None
    
    val = border_el.attrib.get(f"{{{NS['w']}}}val", "none")
    if val == "nil":
        val = "none"
    
    sz = border_el.attrib.get(f"{{{NS['w']}}}sz", "0")
    try:
        width = int(sz)
    except ValueError:
        width = 0
    
    color = border_el.attrib.get(f"{{{NS['w']}}}color")
    if color and color.lower() == "auto":
        color = None
    
    return CellBorder(style=val, width=width, color=color)


def _get_cell_borders(tc_el: ET.Element) -> CellBorders | None:
    """Extract cell border styles."""
    tc_pr = tc_el.find("w:tcPr", NS)
    if tc_pr is None:
        return None
    
    tc_borders = tc_pr.find("w:tcBorders", NS)
    if tc_borders is None:
        return None
    
    return CellBorders(
        top=_parse_border(tc_borders.find("w:top", NS)),
        bottom=_parse_border(tc_borders.find("w:bottom", NS)),
        left=_parse_border(tc_borders.find("w:left", NS)),
        right=_parse_border(tc_borders.find("w:right", NS)),
    )


def _get_cell_spans(tc_el: ET.Element) -> tuple[int, int, str | None]:
    """Extract colspan, rowspan, and vMerge from cell properties.
    
    Returns: (col_span, row_span, v_merge)
    - col_span: Number of columns this cell spans (from gridSpan)
    - row_span: Always 1 for now (vMerge is handled separately)
    - v_merge: "restart" if this starts a vertical merge, "continue" if continuation, None otherwise
    """
    tc_pr = tc_el.find("w:tcPr", NS)
    if tc_pr is None:
        return 1, 1, None
    
    # Horizontal span (gridSpan)
    col_span = 1
    grid_span = tc_pr.find("w:gridSpan", NS)
    if grid_span is not None:
        try:
            col_span = int(grid_span.attrib.get(f"{{{NS['w']}}}val", "1"))
        except ValueError:
            col_span = 1
    
    # Vertical merge
    v_merge = None
    v_merge_el = tc_pr.find("w:vMerge", NS)
    if v_merge_el is not None:
        # If val is present and is "restart", this starts the merge
        # If val is absent or "continue", this is a continuation
        v_merge = v_merge_el.attrib.get(f"{{{NS['w']}}}val", "continue")
    
    return col_span, 1, v_merge


def _extract_sdt_info(sdt_el: ET.Element) -> dict | None:
    """Extract SDT (content control) info - returns type and properties.
    
    Returns dict with 'type' = 'checkbox' | 'dropdown' | 'text' | None
    and relevant properties, or None if not a recognized SDT type.
    """
    sdt_pr = sdt_el.find("w:sdtPr", NS)
    if sdt_pr is None:
        return None
    
    # Check for checkbox (w14:checkbox)
    checkbox_el = sdt_pr.find("w14:checkbox", NS)
    if checkbox_el is not None:
        alias_el = sdt_pr.find("w:alias", NS)
        tag_el = sdt_pr.find("w:tag", NS)
        id_el = sdt_pr.find("w:id", NS)
        checked_el = checkbox_el.find("w14:checked", NS)
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "") if id_el is not None else ""
        label = alias_el.attrib.get(f"{{{NS['w']}}}val", "") if alias_el is not None else ""
        tag = tag_el.attrib.get(f"{{{NS['w']}}}val", "") if tag_el is not None else ""
        
        checked = False
        if checked_el is not None:
            checked_val = checked_el.attrib.get(f"{{{NS['w14']}}}val", "0")
            checked = checked_val == "1"
        
        return {
            "type": "checkbox",
            "id": f"checkbox-{field_id}",
            "xml_ref": f"sdt[id={field_id}]",
            "label": label or tag or f"Checkbox {field_id}",
            "checked": checked,
        }
    
    # Check for dropdown (w:dropDownList) or combobox (w:comboBox)
    dropdown_el = sdt_pr.find("w:dropDownList", NS)
    if dropdown_el is None:
        dropdown_el = sdt_pr.find("w:comboBox", NS)
    
    if dropdown_el is not None:
        alias_el = sdt_pr.find("w:alias", NS)
        tag_el = sdt_pr.find("w:tag", NS)
        id_el = sdt_pr.find("w:id", NS)
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "") if id_el is not None else ""
        label = alias_el.attrib.get(f"{{{NS['w']}}}val", "") if alias_el is not None else ""
        tag = tag_el.attrib.get(f"{{{NS['w']}}}val", "") if tag_el is not None else ""
        
        # Extract options
        options = []
        for item in dropdown_el.findall("w:listItem", NS):
            display = item.attrib.get(f"{{{NS['w']}}}displayText", "")
            value = item.attrib.get(f"{{{NS['w']}}}value", "")
            options.append(display or value)
        
        # Get selected value from content
        selected = None
        sdt_content = sdt_el.find("w:sdtContent", NS)
        if sdt_content is not None:
            for t_el in sdt_content.iter(f"{{{NS['w']}}}t"):
                if t_el.text:
                    selected = t_el.text.strip()
                    break
        
        return {
            "type": "dropdown",
            "id": f"dropdown-{field_id}",
            "xml_ref": f"sdt[id={field_id}]",
            "label": label or tag or f"Dropdown {field_id}",
            "options": options,
            "selected": selected,
        }
    
    # Plain text or rich text SDT - just extract text runs from content
    return {"type": "text_sdt"}


def _paragraph_to_block(p_el: ET.Element, p_index: int) -> ParagraphBlock:
    """Convert a paragraph element to a ParagraphBlock with inline content.
    
    Inline content can be:
    - TextRun: Regular text runs
    - CheckboxRun: Checkbox content controls
    - DropdownRun: Dropdown/combo content controls
    """
    inline_content: List[InlineContent] = []
    run_index = 0
    sdt_index = 0
    
    def process_children(parent_el: ET.Element):
        """Process children, handling runs and SDTs."""
        nonlocal run_index, sdt_index
        
        for child in parent_el:
            if child.tag == f"{{{NS['w']}}}r":
                # Regular text run
                xml_ref = f"p[{p_index}]/r[{run_index}]"
                inline_content.append(
                    TextRun(
                        id=f"run-{p_index}-{run_index}",
                        xml_ref=xml_ref,
                        text=_get_text_from_run(child),
                        bold=_is_bold(child),
                        italic=_is_italic(child),
                        color=_get_text_color(child),
                    )
                )
                run_index += 1
                
            elif child.tag == f"{{{NS['w']}}}sdt":
                # SDT content control - check type
                sdt_info = _extract_sdt_info(child)
                
                if sdt_info is None:
                    continue
                    
                if sdt_info["type"] == "checkbox":
                    inline_content.append(
                        CheckboxRun(
                            id=sdt_info["id"],
                            xml_ref=sdt_info["xml_ref"],
                            label=sdt_info["label"],
                            checked=sdt_info["checked"],
                        )
                    )
                    sdt_index += 1
                    
                elif sdt_info["type"] == "dropdown":
                    inline_content.append(
                        DropdownRun(
                            id=sdt_info["id"],
                            xml_ref=sdt_info["xml_ref"],
                            label=sdt_info["label"],
                            options=sdt_info["options"],
                            selected=sdt_info["selected"],
                        )
                    )
                    sdt_index += 1
                    
                elif sdt_info["type"] == "text_sdt":
                    # Plain text SDT - extract runs from content
                    sdt_content = child.find("w:sdtContent", NS)
                    if sdt_content is not None:
                        # Recursively process SDT content
                        process_children(sdt_content)
            
            elif child.tag == f"{{{NS['w']}}}hyperlink":
                # Hyperlinks can contain runs - process them
                process_children(child)
    
    # Process direct children of the paragraph
    process_children(p_el)

    ppr = p_el.find("w:pPr", NS)
    style_name = None
    if ppr is not None:
        p_style = ppr.find("w:pStyle", NS)
        if p_style is not None:
            style_name = p_style.attrib.get(f"{{{NS['w']}}}val")

    return ParagraphBlock(
        id=f"p-{p_index}",
        xml_ref=f"p[{p_index}]",
        style_name=style_name,
        runs=inline_content,
    )


def _iter_table_rows(tbl_el):
    """Yield all table rows, including those wrapped in SDT content controls."""
    for child in tbl_el:
        if child.tag == f"{{{NS['w']}}}tr":
            yield child
        elif child.tag == f"{{{NS['w']}}}sdt":
            # SDT at table level - unwrap and find rows inside
            sdt_content = child.find("w:sdtContent", NS)
            if sdt_content is not None:
                # Recursively process SDT content (may have nested SDTs)
                yield from _iter_table_rows(sdt_content)


def _iter_row_cells(tr_el):
    """Yield (cell, sdt_info) tuples for cells in a table row.
    
    Returns tuples of (tc_element, sdt_info_dict_or_None).
    If the cell is wrapped in a checkbox/dropdown SDT, sdt_info contains the control info.
    This only yields cells that are direct children of the row,
    NOT cells from nested tables within those cells.
    """
    for child in tr_el:
        if child.tag == f"{{{NS['w']}}}tc":
            yield (child, None)
        elif child.tag == f"{{{NS['w']}}}sdt":
            # SDT wrapping a cell - check if it's a control type
            sdt_info = _extract_sdt_info(child)
            sdt_content = child.find("w:sdtContent", NS)
            if sdt_content is not None:
                # Yield cells from inside the SDT, passing along the SDT info
                for tc in sdt_content:
                    if tc.tag == f"{{{NS['w']}}}tc":
                        yield (tc, sdt_info)
                    elif tc.tag == f"{{{NS['w']}}}sdt":
                        # Nested SDT - recurse
                        for nested_result in _iter_row_cells(sdt_content):
                            yield nested_result


def _iter_cell_paragraphs(tc_el):
    """Yield paragraphs in a cell, handling SDT wrapping.
    
    This handles the case where paragraphs are wrapped in SDT content controls.
    Only yields direct paragraphs and SDT-wrapped paragraphs, NOT paragraphs
    from nested tables.
    """
    for child in tc_el:
        if child.tag == f"{{{NS['w']}}}p":
            yield child
        elif child.tag == f"{{{NS['w']}}}sdt":
            # SDT might contain paragraphs - unwrap them
            sdt_content = child.find("w:sdtContent", NS)
            if sdt_content is not None:
                for sdt_child in sdt_content:
                    if sdt_child.tag == f"{{{NS['w']}}}p":
                        yield sdt_child


def _table_to_block(
    tbl_el: ET.Element, 
    tbl_index: int, 
    nested_tbl_counter: List[int] = None,
    parent_ref_prefix: str = ""
) -> TableBlock:
    """Parse a table element into a TableBlock, including nested tables.
    
    Args:
        tbl_el: The w:tbl XML element
        tbl_index: Index of this table at its level
        nested_tbl_counter: Mutable counter for nested tables (for unique IDs)
        parent_ref_prefix: Prefix for xml_ref paths (for nested tables)
    """
    if nested_tbl_counter is None:
        nested_tbl_counter = [0]  # Use list for mutable counter
    
    # Build the base ref for this table
    if parent_ref_prefix:
        tbl_ref = f"{parent_ref_prefix}/tbl[{tbl_index}]"
    else:
        tbl_ref = f"tbl[{tbl_index}]"
    
    rows: List[TableRow] = []
    row_index = 0
    for tr_el in _iter_table_rows(tbl_el):
        cells: List[TableCell] = []
        cell_index = 0
        # Find direct child cells only (not from nested tables)
        # Handle SDT-wrapped cells by iterating children
        for tc_el, row_level_sdt_info in _iter_row_cells(tr_el):
            # Extract spans and merge info
            col_span, row_span, v_merge = _get_cell_spans(tc_el)

            # Build cell ref prefix for nested content
            cell_ref = f"{tbl_ref}/tr[{row_index}]/tc[{cell_index}]"
            
            # Extract cell content: paragraphs and nested tables
            cell_blocks = []
            cell_p_index = 0
            cell_nested_tbl_index = 0
            
            # If the cell was wrapped in a row-level control SDT, add that control first
            if row_level_sdt_info and row_level_sdt_info["type"] == "checkbox":
                para_block = ParagraphBlock(
                    id=f"p-{cell_p_index}",
                    xml_ref=f"{cell_ref}/p[{cell_p_index}]",
                    style_name=None,
                    runs=[CheckboxRun(
                        id=row_level_sdt_info["id"],
                        xml_ref=row_level_sdt_info["xml_ref"],
                        label=row_level_sdt_info["label"],
                        checked=row_level_sdt_info["checked"],
                    )],
                )
                cell_blocks.append(para_block)
                cell_p_index += 1
            elif row_level_sdt_info and row_level_sdt_info["type"] == "dropdown":
                para_block = ParagraphBlock(
                    id=f"p-{cell_p_index}",
                    xml_ref=f"{cell_ref}/p[{cell_p_index}]",
                    style_name=None,
                    runs=[DropdownRun(
                        id=row_level_sdt_info["id"],
                        xml_ref=row_level_sdt_info["xml_ref"],
                        label=row_level_sdt_info["label"],
                        options=row_level_sdt_info["options"],
                        selected=row_level_sdt_info["selected"],
                    )],
                )
                cell_blocks.append(para_block)
                cell_p_index += 1
            
            # Process direct children of the cell to maintain order and avoid
            # capturing content from nested tables
            for child in tc_el:
                if child.tag == f"{{{NS['w']}}}p":
                    # Direct paragraph in cell
                    para_block = _paragraph_to_block(child, cell_p_index)
                    para_block.xml_ref = f"{cell_ref}/{para_block.xml_ref}"
                    cell_blocks.append(para_block)
                    cell_p_index += 1
                elif child.tag == f"{{{NS['w']}}}tbl":
                    # Nested table in cell - parse recursively with full path
                    nested_tbl_counter[0] += 1
                    nested_tbl = _table_to_block(
                        child, 
                        cell_nested_tbl_index,  # Use local index within cell
                        nested_tbl_counter,
                        parent_ref_prefix=cell_ref  # Pass full path as prefix
                    )
                    nested_tbl.id = f"nested-tbl-{nested_tbl_counter[0]}"
                    cell_blocks.append(nested_tbl)
                    cell_nested_tbl_index += 1
                elif child.tag == f"{{{NS['w']}}}sdt":
                    # SDT at cell level - check if it's a checkbox/dropdown
                    sdt_info = _extract_sdt_info(child)
                    
                    if sdt_info and sdt_info["type"] == "checkbox":
                        # Create a paragraph with the checkbox as inline content
                        para_block = ParagraphBlock(
                            id=f"p-{cell_p_index}",
                            xml_ref=f"{cell_ref}/p[{cell_p_index}]",
                            style_name=None,
                            runs=[CheckboxRun(
                                id=sdt_info["id"],
                                xml_ref=sdt_info["xml_ref"],
                                label=sdt_info["label"],
                                checked=sdt_info["checked"],
                            )],
                        )
                        cell_blocks.append(para_block)
                        cell_p_index += 1
                        
                    elif sdt_info and sdt_info["type"] == "dropdown":
                        # Create a paragraph with the dropdown as inline content
                        para_block = ParagraphBlock(
                            id=f"p-{cell_p_index}",
                            xml_ref=f"{cell_ref}/p[{cell_p_index}]",
                            style_name=None,
                            runs=[DropdownRun(
                                id=sdt_info["id"],
                                xml_ref=sdt_info["xml_ref"],
                                label=sdt_info["label"],
                                options=sdt_info["options"],
                                selected=sdt_info["selected"],
                            )],
                        )
                        cell_blocks.append(para_block)
                        cell_p_index += 1
                        
                    else:
                        # Plain text SDT - extract paragraphs from content
                        sdt_content = child.find("w:sdtContent", NS)
                        if sdt_content is not None:
                            for sdt_child in sdt_content:
                                if sdt_child.tag == f"{{{NS['w']}}}p":
                                    para_block = _paragraph_to_block(sdt_child, cell_p_index)
                                    para_block.xml_ref = f"{cell_ref}/{para_block.xml_ref}"
                                    cell_blocks.append(para_block)
                                    cell_p_index += 1

            cells.append(
                TableCell(
                    id=f"cell-{nested_tbl_counter[0] if parent_ref_prefix else tbl_index}-{row_index}-{cell_index}",
                    xml_ref=cell_ref,
                    row_span=row_span,
                    col_span=col_span,
                    background_color=_get_cell_background(tc_el),
                    borders=_get_cell_borders(tc_el),
                    v_merge=v_merge,
                    blocks=cell_blocks,
                )
            )
            cell_index += 1

        row_xml_ref = f"{tbl_ref}/tr[{row_index}]"
        rows.append(
            TableRow(
                id=f"row-{nested_tbl_counter[0] if parent_ref_prefix else tbl_index}-{row_index}",
                xml_ref=row_xml_ref,
                cells=cells,
            )
        )
        row_index += 1

    return TableBlock(
        id=f"tbl-{tbl_index}" if not parent_ref_prefix else f"nested-tbl-{nested_tbl_counter[0]}",
        xml_ref=tbl_ref,
        rows=rows,
    )


def docx_to_json(docx_path: str, document_id: str) -> DocumentJSON:
    """Parse a DOCX into our DocumentJSON structure.

    v2: paragraphs, tables, and drawings, using positional xml_ref paths.
    """

    docx_path = str(docx_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        with zf.open("word/document.xml") as doc_xml:
            tree = ET.parse(doc_xml)

    root = tree.getroot()
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("Invalid DOCX: missing w:body")

    blocks: List[Block] = []

    p_counter = 0
    tbl_counter = 0
    drawing_counter = 0
    for kind, el in _iter_body_elements(body):
        if kind == "p":
            block = _paragraph_to_block(el, p_counter)
            blocks.append(block)
            p_counter += 1
        elif kind == "tbl":
            block = _table_to_block(el, tbl_counter)
            blocks.append(block)
            tbl_counter += 1
        elif kind == "drawing":
            drawing_block = _extract_drawing(el, drawing_counter)
            if drawing_block:
                blocks.append(drawing_block)
            drawing_counter += 1

    # Extract form fields (checkboxes and dropdowns)
    checkboxes = _extract_checkboxes(root)
    dropdowns = _extract_dropdowns(root)

    return DocumentJSON(
        id=document_id,
        title=None,
        blocks=blocks,
        checkboxes=checkboxes,
        dropdowns=dropdowns,
    )


# =============================================================================
# EXPORT: JSON → DOCX
# =============================================================================

def _iter_body_children_by_tag(body_el: ET.Element, target_tag: str):
    """Iterate body children matching target_tag, unwrapping SDT elements.
    
    This mirrors the logic in _iter_body_elements to ensure consistent indexing.
    """
    for el in body_el:
        if el.tag == target_tag:
            yield el
        elif el.tag == f"{{{NS['w']}}}sdt":
            # SDT at body level - unwrap and check contents
            sdt_content = el.find("w:sdtContent", NS)
            if sdt_content is not None:
                yield from _iter_body_children_by_tag(sdt_content, target_tag)


def _find_node_by_ref(body_el: ET.Element, xml_ref: str) -> ET.Element | None:
    """Resolve a simple positional xml_ref into an XML element.

    Supported patterns:
    - p[i]
    - p[i]/r[j]
    - tbl[i]
    - tbl[i]/tr[j]
    - tbl[i]/tr[j]/tc[k]
    - tbl[i]/tr[j]/tc[k]/p[m]
    - tbl[i]/tr[j]/tc[k]/tbl[n]/... (nested tables)
    - ... and nested /r[m]
    
    IMPORTANT: Body-level elements use SDT-unwrapping iteration to match
    the parsing logic in _iter_body_elements. Elements inside table cells
    use appropriate iteration helpers to handle SDT wrapping.
    """
    parts = xml_ref.split("/")
    current: ET.Element = body_el
    is_inside_cell = False  # Track if we're inside a table cell
    is_body_level = True  # Track if we're still at body level

    tag_map = {
        "p": f"{{{NS['w']}}}p",
        "tbl": f"{{{NS['w']}}}tbl",
        "tr": f"{{{NS['w']}}}tr",
        "tc": f"{{{NS['w']}}}tc",
        "r": f"{{{NS['w']}}}r",
    }

    for i, part in enumerate(parts):
        m = re.match(r"(\w+)\[(\d+)\]", part)
        if not m:
            return None
        
        name, idx_s = m.group(1), m.group(2)
        idx = int(idx_s)
        
        xml_tag = tag_map.get(name)
        if xml_tag is None:
            return None
        
        # Determine search strategy based on CURRENT context:
        if is_body_level and name in ("p", "tbl"):
            # Body-level: use SDT-unwrapping iteration to match parsing
            matches = list(_iter_body_children_by_tag(current, xml_tag))
        elif name == "tr":
            # Table rows: use helper to handle SDT-wrapped rows
            matches = list(_iter_table_rows(current))
        elif name == "tc":
            # Table cells: use helper to handle SDT-wrapped cells
            # _iter_row_cells returns (cell, sdt_info) tuples - extract just cells
            matches = [cell for cell, _ in _iter_row_cells(current)]
        elif is_inside_cell and name == "tbl":
            # Nested table inside a cell: direct children only
            matches = [child for child in current if child.tag == xml_tag]
        elif is_inside_cell and name == "p":
            # Paragraphs inside cells: use helper to handle SDT-wrapped paragraphs
            matches = list(_iter_cell_paragraphs(current))
        elif is_inside_cell and name == "r":
            # Runs inside paragraphs: direct children (runs aren't SDT-wrapped)
            matches = [child for child in current if child.tag == xml_tag]
        else:
            # Direct children only
            matches = [child for child in current if child.tag == xml_tag]
        
        # Update state AFTER search strategy is determined
        if name == "tbl":
            is_body_level = False
            is_inside_cell = False  # Reset when entering a table
        elif name == "tc":
            is_body_level = False
            is_inside_cell = True  # Now inside a cell
        elif name == "tr":
            is_body_level = False
        
        if 0 <= idx < len(matches):
            current = matches[idx]
        else:
            return None

    return current


def _apply_checkbox_changes(root: ET.Element, checkboxes: List[CheckboxField]) -> None:
    """Apply checkbox state changes to the XML tree.
    
    For each checkbox, we need to:
    1. Update w14:checked val attribute (0 or 1)
    2. Update the display character (☐ U+2610 unchecked, ☒ U+2612 checked)
    """
    # Build a map of checkbox id -> checked state
    checkbox_map = {cb.id: cb.checked for cb in checkboxes}
    
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is None:
            continue
        
        checkbox_el = sdt_pr.find("w14:checkbox", NS)
        if checkbox_el is None:
            continue
        
        id_el = sdt_pr.find("w:id", NS)
        if id_el is None:
            continue
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "")
        checkbox_id = f"checkbox-{field_id}"
        
        if checkbox_id not in checkbox_map:
            continue
        
        is_checked = checkbox_map[checkbox_id]
        
        # Update w14:checked element
        checked_el = checkbox_el.find("w14:checked", NS)
        if checked_el is not None:
            checked_el.set(f"{{{NS['w14']}}}val", "1" if is_checked else "0")
        
        # Update display text (the checkbox symbol)
        sdt_content = sdt.find("w:sdtContent", NS)
        if sdt_content is not None:
            for t_el in sdt_content.iter(f"{{{NS['w']}}}t"):
                # Replace checkbox symbol
                t_el.text = "\u2612" if is_checked else "\u2610"
                break


def _apply_dropdown_changes(root: ET.Element, dropdowns: List[DropdownField]) -> None:
    """Apply dropdown selection changes to the XML tree.
    
    For each dropdown, update the text content in sdtContent to the selected value.
    """
    # Build a map of dropdown id -> selected value
    dropdown_map = {dd.id: dd.selected for dd in dropdowns}
    
    for sdt in root.iter(f"{{{NS['w']}}}sdt"):
        sdt_pr = sdt.find("w:sdtPr", NS)
        if sdt_pr is None:
            continue
        
        # Check for comboBox or dropDownList
        combo_el = sdt_pr.find("w:comboBox", NS)
        dropdown_el = sdt_pr.find("w:dropDownList", NS)
        
        if combo_el is None and dropdown_el is None:
            continue
        
        id_el = sdt_pr.find("w:id", NS)
        if id_el is None:
            continue
        
        field_id = id_el.attrib.get(f"{{{NS['w']}}}val", "")
        dropdown_id = f"dropdown-{field_id}"
        
        if dropdown_id not in dropdown_map:
            continue
        
        selected_value = dropdown_map[dropdown_id]
        if selected_value is None:
            continue
        
        # Update the text content
        sdt_content = sdt.find("w:sdtContent", NS)
        if sdt_content is not None:
            for t_el in sdt_content.iter(f"{{{NS['w']}}}t"):
                t_el.text = selected_value
                break


def _patch_paragraph_runs(body: ET.Element, para_block: ParagraphBlock) -> None:
    """Patch text in a paragraph by updating each run individually.
    
    This preserves run-level formatting (bold, italic, color) by updating
    each w:r element's w:t text separately rather than concatenating everything.
    
    For empty cells (no w:r or w:t elements), we create the necessary structure.
    
    Note: Only TextRuns are patched here. CheckboxRun and DropdownRun are handled
    separately by _apply_checkbox_changes and _apply_dropdown_changes.
    """
    # Find the paragraph element using its xml_ref
    p_el = _find_node_by_ref(body, para_block.xml_ref)
    if p_el is None:
        return
    
    # Filter to only TextRuns (skip CheckboxRun and DropdownRun)
    text_runs = [run for run in para_block.runs if isinstance(run, TextRun)]
    
    # Get all w:r elements in the paragraph (including nested in SDT)
    r_els = p_el.findall(".//w:r", NS)
    
    # If no runs exist but we have text to add, create a run
    if not r_els and text_runs:
        # Check if there's text to add
        has_text = any(run.text for run in text_runs)
        if has_text:
            # Create a new w:r element with w:t inside
            for run in text_runs:
                if run.text:
                    new_r = ET.SubElement(p_el, f"{{{NS['w']}}}r")
                    new_t = ET.SubElement(new_r, f"{{{NS['w']}}}t")
                    new_t.text = run.text
                    # Set xml:space="preserve" to keep whitespace
                    new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    
    # Match runs by index - each run in JSON corresponds to a w:r element
    for i, run in enumerate(text_runs):
        if i >= len(r_els):
            # More runs in JSON than in XML - create new run if we have text
            if run.text:
                new_r = ET.SubElement(p_el, f"{{{NS['w']}}}r")
                new_t = ET.SubElement(new_r, f"{{{NS['w']}}}t")
                new_t.text = run.text
                new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            continue
        
        r_el = r_els[i]
        
        # Find w:t elements within this run
        t_els = r_el.findall("w:t", NS)
        if not t_els:
            # No text element exists - create one if we have text to add
            if run.text:
                new_t = ET.SubElement(r_el, f"{{{NS['w']}}}t")
                new_t.text = run.text
                new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            continue
        
        # Update the first w:t with the run's text
        t_els[0].text = run.text or ""
        
        # Clear any additional w:t elements in this run (rare but possible)
        for t_el in t_els[1:]:
            t_el.text = ""


def apply_json_to_docx(json_doc: DocumentJSON, base_docx_path: str, out_docx_path: str) -> str:
    """Emit a *new* DOCX file with text edits applied from DocumentJSON.

    Strategy: For each run in the JSON, find the corresponding w:r element
    and update its w:t text. This preserves run-level formatting.
    All other parts of the DOCX are copied byte-for-byte from the base file.
    """

    base = Path(base_docx_path)
    out = Path(out_docx_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    # Read and patch word/document.xml in-memory
    with zipfile.ZipFile(base, "r") as zin:
        with zin.open("word/document.xml") as doc_xml:
            tree = ET.parse(doc_xml)

    root = tree.getroot()
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("Invalid DOCX: missing w:body")

    # Patch paragraph text according to json_doc
    # Strategy: For each run in the JSON, find the corresponding w:r element
    # and update its w:t text. This preserves run-level formatting.
    
    def patch_table_blocks(table_block: TableBlock):
        """Recursively patch paragraphs in a table and its nested tables."""
        for row in table_block.rows:
            for cell in row.cells:
                for cell_block in cell.blocks:
                    if isinstance(cell_block, ParagraphBlock):
                        _patch_paragraph_runs(body, cell_block)
                    elif isinstance(cell_block, TableBlock):
                        # Nested table - patch recursively
                        patch_table_blocks(cell_block)
    
    for block in json_doc.blocks:
        if isinstance(block, ParagraphBlock):
            _patch_paragraph_runs(body, block)
        elif isinstance(block, TableBlock):
            patch_table_blocks(block)
        # DrawingBlocks are not patched - they're preserved as-is from the base DOCX

    # Collect inline checkboxes/dropdowns from blocks for applying changes
    def collect_inline_controls(blocks):
        """Recursively collect inline CheckboxRun and DropdownRun from blocks."""
        inline_checkboxes = []
        inline_dropdowns = []
        
        for block in blocks:
            if isinstance(block, ParagraphBlock):
                for run in block.runs:
                    if isinstance(run, CheckboxRun):
                        inline_checkboxes.append(CheckboxField(
                            id=run.id,
                            xml_ref=run.xml_ref,
                            label=run.label,
                            checked=run.checked,
                        ))
                    elif isinstance(run, DropdownRun):
                        inline_dropdowns.append(DropdownField(
                            id=run.id,
                            xml_ref=run.xml_ref,
                            label=run.label,
                            options=run.options,
                            selected=run.selected,
                        ))
            elif isinstance(block, TableBlock):
                for row in block.rows:
                    for cell in row.cells:
                        cb, dd = collect_inline_controls(cell.blocks)
                        inline_checkboxes.extend(cb)
                        inline_dropdowns.extend(dd)
        
        return inline_checkboxes, inline_dropdowns
    
    inline_cb, inline_dd = collect_inline_controls(json_doc.blocks)
    
    # Merge legacy arrays with inline controls (inline takes precedence by id)
    all_checkboxes = {cb.id: cb for cb in json_doc.checkboxes}
    for cb in inline_cb:
        all_checkboxes[cb.id] = cb  # Inline overrides legacy
    
    all_dropdowns = {dd.id: dd for dd in json_doc.dropdowns}
    for dd in inline_dd:
        all_dropdowns[dd.id] = dd  # Inline overrides legacy
    
    # Patch checkboxes (using merged set)
    _apply_checkbox_changes(root, list(all_checkboxes.values()))
    
    # Patch dropdowns (using merged set)
    _apply_dropdown_changes(root, list(all_dropdowns.values()))

    # Serialize the modified XML tree
    if USING_LXML:
        # lxml preserves namespace prefixes properly and produces Word-compatible output
        new_document_xml = ET.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,  # Word requires standalone="yes"
        )
    else:
        # ElementTree fallback (may cause namespace prefix issues)
        buffer = BytesIO()
        tree.write(buffer, xml_declaration=True, encoding="UTF-8", method="xml")
        new_document_xml = buffer.getvalue()
        
        # Fix XML declaration: Word requires standalone="yes" and double quotes
        # ElementTree produces: <?xml version='1.0' encoding='UTF-8'?>
        # Word expects: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        new_document_xml = new_document_xml.replace(
            b"<?xml version='1.0' encoding='UTF-8'?>",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        )

    # Rebuild DOCX: copy all entries, replacing word/document.xml
    # Use ZIP_DEFLATED compression like the original file
    with zipfile.ZipFile(base, "r") as zin, zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                continue
            data = zin.read(item.filename)
            # Preserve original compression info
            zout.writestr(item, data)

        # Write the modified document.xml with same compression
        zout.writestr("word/document.xml", new_document_xml, compress_type=zipfile.ZIP_DEFLATED)

    return str(out)


# =============================================================================
# VALIDATION
# =============================================================================

def validate_document_json(doc: DocumentJSON) -> ValidationResult:
    """Basic structural validation for a DocumentJSON.

    v1: ensure IDs and xml_refs are populated and unique at block level.
    """

    errors: List[ValidationErrorDetail] = []

    block_ids = set()
    for block in doc.blocks:
        if block.id in block_ids:
            errors.append(
                ValidationErrorDetail(
                    field="blocks.id",
                    message=f"Duplicate block id: {block.id}",
                )
            )
        else:
            block_ids.add(block.id)

        if not block.xml_ref:
            errors.append(
                ValidationErrorDetail(
                    field=f"block[{block.id}].xml_ref",
                    message="xml_ref is required",
                )
            )

        if isinstance(block, ParagraphBlock):
            for run in block.runs:
                if not run.xml_ref:
                    errors.append(
                        ValidationErrorDetail(
                            field=f"run[{run.id}].xml_ref",
                            message="xml_ref is required",
                        )
                    )

        elif isinstance(block, TableBlock):
            for row in block.rows:
                if not row.xml_ref:
                    errors.append(
                        ValidationErrorDetail(
                            field=f"row[{row.id}].xml_ref",
                            message="xml_ref is required",
                        )
                    )
                for cell in row.cells:
                    if not cell.xml_ref:
                        errors.append(
                            ValidationErrorDetail(
                                field=f"cell[{cell.id}].xml_ref",
                                message="xml_ref is required",
                            )
                        )

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)
