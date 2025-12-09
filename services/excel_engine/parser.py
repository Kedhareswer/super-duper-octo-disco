"""XLSX Parser - Converts Excel files to JSON structure.

Handles complex elements:
- Multiple worksheets
- Merged cells
- Data validation (dropdowns)
- Cell styles and formatting
- Images and drawings
- Comments/notes
- Formulas
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .schemas import (
    CellAlignment,
    CellBorder,
    CellBorders,
    CellDataType,
    CellFill,
    CellFont,
    CellStyle,
    ColumnInfo,
    DataValidationRule,
    ExcelCellJSON,
    ExcelComment,
    ExcelImage,
    ExcelSheetJSON,
    ExcelWorkbookJSON,
    MergedCellRange,
    RowInfo,
    SharedStringItem,
    StyleInfo,
    # New complex elements
    ExcelHyperlink,
    ConditionalFormatting,
    ConditionalFormatRule,
    FormControl,
    FormControlType,
    DefinedName,
    SheetView,
    FreezePane,
    ExcelTable,
    TableColumn,
    Sparkline,
    SparklineGroup,
)


# =============================================================================
# NAMESPACES
# =============================================================================

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "vml": "urn:schemas-microsoft-com:vml",
    "x": "urn:schemas-microsoft-com:office:excel",
    "o": "urn:schemas-microsoft-com:office:office",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
}

# Register namespaces
for prefix, uri in NS.items():
    ET.register_namespace(prefix if prefix != "main" else "", uri)


# =============================================================================
# UTILITIES
# =============================================================================

def col_letter_to_index(col: str) -> int:
    """Convert column letter(s) to 1-indexed number. A=1, B=2, ..., Z=26, AA=27."""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def col_index_to_letter(index: int) -> str:
    """Convert 1-indexed column number to letter(s). 1=A, 2=B, ..., 27=AA."""
    result = ""
    while index > 0:
        index -= 1
        result = chr(ord('A') + (index % 26)) + result
        index //= 26
    return result


def parse_cell_ref(ref: str) -> Tuple[str, int, int]:
    """Parse cell reference like 'A1' or 'AA100' into (col_letter, col_num, row_num)."""
    match = re.match(r'^([A-Z]+)(\d+)$', ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    col_letter = match.group(1)
    row = int(match.group(2))
    col = col_letter_to_index(col_letter)
    return col_letter, col, row


def parse_range_ref(ref: str) -> Tuple[int, int, int, int]:
    """Parse range like 'B2:F6' into (start_row, start_col, end_row, end_col)."""
    parts = ref.split(':')
    if len(parts) != 2:
        raise ValueError(f"Invalid range reference: {ref}")
    
    start_letter, start_col, start_row = parse_cell_ref(parts[0])
    end_letter, end_col, end_row = parse_cell_ref(parts[1])
    
    return start_row, start_col, end_row, end_col


# =============================================================================
# SHARED STRINGS
# =============================================================================

def _parse_shared_strings(zf: zipfile.ZipFile) -> List[SharedStringItem]:
    """Parse shared strings table."""
    shared_strings: List[SharedStringItem] = []
    
    try:
        with zf.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()
            
            ns = NS["main"]
            for i, si in enumerate(root.findall(f"{{{ns}}}si")):
                # Simple text
                t_el = si.find(f"{{{ns}}}t")
                if t_el is not None and t_el.text:
                    shared_strings.append(SharedStringItem(
                        index=i,
                        text=t_el.text,
                        rich_text=False
                    ))
                else:
                    # Rich text (multiple <r> elements)
                    text_parts = []
                    for r in si.findall(f"{{{ns}}}r"):
                        t = r.find(f"{{{ns}}}t")
                        if t is not None and t.text:
                            text_parts.append(t.text)
                    shared_strings.append(SharedStringItem(
                        index=i,
                        text="".join(text_parts),
                        rich_text=True
                    ))
    except KeyError:
        pass  # No shared strings
    
    return shared_strings


# =============================================================================
# STYLES
# =============================================================================

def _parse_styles(zf: zipfile.ZipFile) -> StyleInfo:
    """Parse styles.xml for cell formatting."""
    style_info = StyleInfo()
    
    try:
        with zf.open("xl/styles.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = NS["main"]
            
            # Parse fonts
            fonts_el = root.find(f"{{{ns}}}fonts")
            if fonts_el is not None:
                for font_el in fonts_el.findall(f"{{{ns}}}font"):
                    font_dict = {}
                    
                    name_el = font_el.find(f"{{{ns}}}name")
                    if name_el is not None:
                        font_dict["name"] = name_el.get("val")
                    
                    sz_el = font_el.find(f"{{{ns}}}sz")
                    if sz_el is not None:
                        font_dict["size"] = float(sz_el.get("val", 11))
                    
                    if font_el.find(f"{{{ns}}}b") is not None:
                        font_dict["bold"] = True
                    if font_el.find(f"{{{ns}}}i") is not None:
                        font_dict["italic"] = True
                    if font_el.find(f"{{{ns}}}u") is not None:
                        font_dict["underline"] = True
                    if font_el.find(f"{{{ns}}}strike") is not None:
                        font_dict["strike"] = True
                    
                    color_el = font_el.find(f"{{{ns}}}color")
                    if color_el is not None:
                        rgb = color_el.get("rgb")
                        if rgb:
                            font_dict["color"] = rgb
                    
                    style_info.fonts.append(font_dict)
            
            # Parse fills
            fills_el = root.find(f"{{{ns}}}fills")
            if fills_el is not None:
                for fill_el in fills_el.findall(f"{{{ns}}}fill"):
                    fill_dict = {}
                    pattern_el = fill_el.find(f"{{{ns}}}patternFill")
                    if pattern_el is not None:
                        fill_dict["patternType"] = pattern_el.get("patternType")
                        fg = pattern_el.find(f"{{{ns}}}fgColor")
                        bg = pattern_el.find(f"{{{ns}}}bgColor")
                        if fg is not None:
                            fill_dict["fgColor"] = fg.get("rgb") or fg.get("theme")
                        if bg is not None:
                            fill_dict["bgColor"] = bg.get("rgb") or bg.get("theme")
                    style_info.fills.append(fill_dict)
            
            # Parse borders
            borders_el = root.find(f"{{{ns}}}borders")
            if borders_el is not None:
                for border_el in borders_el.findall(f"{{{ns}}}border"):
                    border_dict = {}
                    for side in ["left", "right", "top", "bottom", "diagonal"]:
                        side_el = border_el.find(f"{{{ns}}}{side}")
                        if side_el is not None:
                            style = side_el.get("style")
                            if style:
                                color_el = side_el.find(f"{{{ns}}}color")
                                color = None
                                if color_el is not None:
                                    color = color_el.get("rgb") or color_el.get("indexed")
                                border_dict[side] = {"style": style, "color": color}
                    style_info.borders.append(border_dict)
            
            # Parse cellXfs (cell formats)
            cell_xfs_el = root.find(f"{{{ns}}}cellXfs")
            if cell_xfs_el is not None:
                for xf in cell_xfs_el.findall(f"{{{ns}}}xf"):
                    xf_dict = {
                        "fontId": int(xf.get("fontId", 0)),
                        "fillId": int(xf.get("fillId", 0)),
                        "borderId": int(xf.get("borderId", 0)),
                        "numFmtId": int(xf.get("numFmtId", 0)),
                    }
                    alignment_el = xf.find(f"{{{ns}}}alignment")
                    if alignment_el is not None:
                        xf_dict["alignment"] = {
                            "horizontal": alignment_el.get("horizontal"),
                            "vertical": alignment_el.get("vertical"),
                            "wrapText": alignment_el.get("wrapText") == "1",
                            "textRotation": alignment_el.get("textRotation"),
                        }
                    style_info.cell_xfs.append(xf_dict)
            
            # Parse number formats
            num_fmts_el = root.find(f"{{{ns}}}numFmts")
            if num_fmts_el is not None:
                for num_fmt in num_fmts_el.findall(f"{{{ns}}}numFmt"):
                    fmt_id = int(num_fmt.get("numFmtId", 0))
                    fmt_code = num_fmt.get("formatCode", "")
                    style_info.number_formats[fmt_id] = fmt_code
    
    except KeyError:
        pass  # No styles
    
    return style_info


def _build_cell_style(style_index: int, style_info: StyleInfo) -> Optional[CellStyle]:
    """Build a CellStyle from style index and parsed style info."""
    if style_index < 0 or style_index >= len(style_info.cell_xfs):
        return None
    
    xf = style_info.cell_xfs[style_index]
    style = CellStyle(style_id=style_index)
    
    # Font
    font_id = xf.get("fontId", 0)
    if font_id < len(style_info.fonts):
        font_data = style_info.fonts[font_id]
        style.font = CellFont(
            name=font_data.get("name"),
            size=font_data.get("size"),
            bold=font_data.get("bold", False),
            italic=font_data.get("italic", False),
            underline=font_data.get("underline", False),
            strike=font_data.get("strike", False),
            color=font_data.get("color"),
        )
    
    # Fill
    fill_id = xf.get("fillId", 0)
    if fill_id < len(style_info.fills):
        fill_data = style_info.fills[fill_id]
        style.fill = CellFill(
            pattern_type=fill_data.get("patternType"),
            fg_color=fill_data.get("fgColor"),
            bg_color=fill_data.get("bgColor"),
        )
    
    # Borders
    border_id = xf.get("borderId", 0)
    if border_id < len(style_info.borders):
        border_data = style_info.borders[border_id]
        borders = CellBorders()
        for side in ["left", "right", "top", "bottom", "diagonal"]:
            if side in border_data:
                setattr(borders, side, CellBorder(
                    style=border_data[side].get("style"),
                    color=border_data[side].get("color"),
                ))
        style.borders = borders
    
    # Alignment
    if "alignment" in xf:
        align = xf["alignment"]
        style.alignment = CellAlignment(
            horizontal=align.get("horizontal"),
            vertical=align.get("vertical"),
            wrap_text=align.get("wrapText", False),
            text_rotation=int(align["textRotation"]) if align.get("textRotation") else None,
        )
    
    # Number format
    num_fmt_id = xf.get("numFmtId", 0)
    if num_fmt_id in style_info.number_formats:
        style.number_format = style_info.number_formats[num_fmt_id]
    
    return style


# =============================================================================
# WORKSHEET PARSING
# =============================================================================

def _parse_merged_cells(sheet_el: ET.Element, sheet_id: str) -> List[MergedCellRange]:
    """Parse merged cell ranges from a worksheet."""
    merged_cells: List[MergedCellRange] = []
    ns = NS["main"]
    
    merge_cells_el = sheet_el.find(f"{{{ns}}}mergeCells")
    if merge_cells_el is None:
        return merged_cells
    
    for i, merge_cell in enumerate(merge_cells_el.findall(f"{{{ns}}}mergeCell")):
        ref = merge_cell.get("ref")
        if not ref:
            continue
        
        try:
            start_row, start_col, end_row, end_col = parse_range_ref(ref)
            start_cell = f"{col_index_to_letter(start_col)}{start_row}"
            end_cell = f"{col_index_to_letter(end_col)}{end_row}"
            
            merged_cells.append(MergedCellRange(
                id=f"{sheet_id}-merge-{i}",
                ref=ref,
                start_row=start_row,
                start_col=start_col,
                end_row=end_row,
                end_col=end_col,
                start_cell_ref=start_cell,
                end_cell_ref=end_cell,
            ))
        except ValueError:
            continue
    
    return merged_cells


def _parse_data_validations(sheet_el: ET.Element, sheet_id: str) -> List[DataValidationRule]:
    """Parse data validation rules (dropdowns, etc.) from a worksheet."""
    validations: List[DataValidationRule] = []
    ns = NS["main"]
    
    dv_el = sheet_el.find(f"{{{ns}}}dataValidations")
    if dv_el is None:
        return validations
    
    for i, dv in enumerate(dv_el.findall(f"{{{ns}}}dataValidation")):
        sqref = dv.get("sqref", "")
        val_type = dv.get("type", "none")
        
        formula1_el = dv.find(f"{{{ns}}}formula1")
        formula2_el = dv.find(f"{{{ns}}}formula2")
        
        formula1 = formula1_el.text if formula1_el is not None and formula1_el.text else None
        formula2 = formula2_el.text if formula2_el is not None and formula2_el.text else None
        
        # Parse options for list type
        options: List[str] = []
        if val_type == "list" and formula1:
            # Remove quotes and split by comma
            clean_formula = formula1.strip('"')
            if clean_formula.startswith("="):
                # Named range or formula reference
                options = [clean_formula]
            else:
                options = [opt.strip() for opt in clean_formula.split(",")]
        
        validations.append(DataValidationRule(
            id=f"{sheet_id}-dv-{i}",
            sqref=sqref,
            validation_type=val_type,
            formula1=formula1,
            formula2=formula2,
            allow_blank=dv.get("allowBlank") == "1",
            show_input_message=dv.get("showInputMessage") == "1",
            show_error_message=dv.get("showErrorMessage") == "1",
            input_title=dv.get("promptTitle"),
            input_message=dv.get("prompt"),
            error_title=dv.get("errorTitle"),
            error_message=dv.get("error"),
            error_style=dv.get("errorStyle"),
            operator=dv.get("operator"),
            options=options,
        ))
    
    return validations


def _parse_columns(sheet_el: ET.Element) -> List[ColumnInfo]:
    """Parse column information from a worksheet."""
    columns: List[ColumnInfo] = []
    ns = NS["main"]
    
    cols_el = sheet_el.find(f"{{{ns}}}cols")
    if cols_el is None:
        return columns
    
    for col in cols_el.findall(f"{{{ns}}}col"):
        min_col = int(col.get("min", 1))
        max_col = int(col.get("max", 1))
        width = float(col.get("width", 0)) if col.get("width") else None
        
        columns.append(ColumnInfo(
            min_col=min_col,
            max_col=max_col,
            width=width,
            hidden=col.get("hidden") == "1",
            best_fit=col.get("bestFit") == "1",
            custom_width=col.get("customWidth") == "1",
            style_index=int(col.get("style", 0)) if col.get("style") else None,
        ))
    
    return columns


def _parse_cells(
    sheet_el: ET.Element,
    sheet_id: str,
    shared_strings: List[SharedStringItem],
    style_info: StyleInfo,
    merged_ranges: List[MergedCellRange],
) -> Tuple[List[ExcelCellJSON], List[RowInfo]]:
    """Parse cells and row info from a worksheet."""
    cells: List[ExcelCellJSON] = []
    rows: List[RowInfo] = []
    ns = NS["main"]
    
    # Build merge lookup: cell_ref -> merge_range
    merge_map: Dict[str, MergedCellRange] = {}
    for merge in merged_ranges:
        for r in range(merge.start_row, merge.end_row + 1):
            for c in range(merge.start_col, merge.end_col + 1):
                ref = f"{col_index_to_letter(c)}{r}"
                merge_map[ref] = merge
    
    sheet_data = sheet_el.find(f"{{{ns}}}sheetData")
    if sheet_data is None:
        return cells, rows
    
    for row_el in sheet_data.findall(f"{{{ns}}}row"):
        row_num = int(row_el.get("r", 0))
        
        # Capture row info
        height = float(row_el.get("ht", 0)) if row_el.get("ht") else None
        if height or row_el.get("hidden") or row_el.get("customHeight"):
            rows.append(RowInfo(
                row=row_num,
                height=height,
                hidden=row_el.get("hidden") == "1",
                custom_height=row_el.get("customHeight") == "1",
                style_index=int(row_el.get("s", 0)) if row_el.get("s") else None,
            ))
        
        for cell_el in row_el.findall(f"{{{ns}}}c"):
            cell_ref = cell_el.get("r")
            if not cell_ref:
                continue
            
            try:
                col_letter, col_num, row_num_parsed = parse_cell_ref(cell_ref)
            except ValueError:
                continue
            
            # Data type
            data_type_str = cell_el.get("t")
            data_type = None
            if data_type_str:
                try:
                    data_type = CellDataType(data_type_str)
                except ValueError:
                    pass
            
            # Value
            v_el = cell_el.find(f"{{{ns}}}v")
            raw_value = v_el.text if v_el is not None else None
            
            # Resolve value
            value: Any = None
            if raw_value is not None:
                if data_type == CellDataType.STRING:
                    # Shared string reference
                    try:
                        ss_index = int(raw_value)
                        if ss_index < len(shared_strings):
                            value = shared_strings[ss_index].text
                        else:
                            value = raw_value
                    except ValueError:
                        value = raw_value
                elif data_type == CellDataType.BOOLEAN:
                    value = raw_value == "1"
                elif data_type == CellDataType.INLINE_STRING:
                    # Inline string
                    is_el = cell_el.find(f"{{{ns}}}is")
                    if is_el is not None:
                        t_el = is_el.find(f"{{{ns}}}t")
                        value = t_el.text if t_el is not None else ""
                else:
                    # Try to parse as number
                    try:
                        if "." in raw_value or "E" in raw_value.upper():
                            value = float(raw_value)
                        else:
                            value = int(raw_value)
                    except ValueError:
                        value = raw_value
            
            # Formula
            f_el = cell_el.find(f"{{{ns}}}f")
            formula = None
            formula_type = None
            shared_formula_ref = None
            shared_formula_si = None
            
            if f_el is not None:
                formula = f_el.text
                formula_type = f_el.get("t", "normal")
                if formula_type == "shared":
                    shared_formula_ref = f_el.get("ref")
                    si = f_el.get("si")
                    if si:
                        shared_formula_si = int(si)
            
            # Style
            style_index = int(cell_el.get("s", 0)) if cell_el.get("s") else None
            style = _build_cell_style(style_index, style_info) if style_index else None
            
            # Merge info
            merge_range = merge_map.get(cell_ref)
            is_merged = merge_range is not None
            is_merge_origin = merge_range.start_cell_ref == cell_ref if merge_range else False
            
            cells.append(ExcelCellJSON(
                id=f"{sheet_id}-{cell_ref}",
                ref=cell_ref,
                row=row_num_parsed,
                col=col_num,
                col_letter=col_letter,
                value=value,
                raw_value=raw_value,
                data_type=data_type,
                formula=formula,
                formula_type=formula_type,
                shared_formula_ref=shared_formula_ref,
                shared_formula_si=shared_formula_si,
                style=style,
                style_index=style_index,
                is_merged=is_merged,
                merge_range=merge_range.ref if merge_range else None,
                is_merge_origin=is_merge_origin,
            ))
    
    return cells, rows


def _parse_images(zf: zipfile.ZipFile, sheet_path: str, sheet_id: str) -> List[ExcelImage]:
    """Parse images from a worksheet's drawing."""
    images: List[ExcelImage] = []
    
    # Find the drawing relationship
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    
    try:
        with zf.open(rels_path) as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
    except KeyError:
        return images
    
    # Find drawing relationship
    drawing_path = None
    ns_rel = NS["rel"]
    for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "drawing" in rel_type.lower():
            target = rel.get("Target", "")
            # Skip VML files - they contain HTML-style content that isn't valid XML
            # VML files are used for form controls, not images we need to extract
            if target.lower().endswith('.vml'):
                continue
            # Resolve relative path
            if target.startswith("../"):
                drawing_path = "xl/" + target[3:]
            else:
                drawing_path = target
            break
    
    if not drawing_path:
        return images
    
    try:
        with zf.open(drawing_path) as f:
            drawing_tree = ET.parse(f)
            drawing_root = drawing_tree.getroot()
    except KeyError:
        return images
    
    # Parse drawing for images
    xdr = NS["xdr"]
    a_ns = NS["a"]
    pic_ns = NS["pic"]
    r_ns = NS["r"]
    
    for i, anchor in enumerate(drawing_root):
        if "Anchor" not in anchor.tag:
            continue
        
        # Get anchor type
        anchor_type = anchor.tag.split("}")[-1] if "}" in anchor.tag else anchor.tag
        
        # Find picture element
        pic = anchor.find(f".//{{{xdr}}}pic")
        if pic is None:
            continue
        
        # Get image properties
        nv_pic_pr = pic.find(f"{{{xdr}}}nvPicPr")
        c_nv_pr = nv_pic_pr.find(f"{{{xdr}}}cNvPr") if nv_pic_pr is not None else None
        
        name = c_nv_pr.get("name") if c_nv_pr is not None else None
        description = c_nv_pr.get("descr") if c_nv_pr is not None else None
        
        # Get the embed relationship
        blip_fill = pic.find(f"{{{xdr}}}blipFill")
        if blip_fill is None:
            continue
        
        blip = blip_fill.find(f"{{{a_ns}}}blip")
        if blip is None:
            continue
        
        r_embed = blip.get(f"{{{r_ns}}}embed")
        
        # Get position
        from_el = anchor.find(f"{{{xdr}}}from")
        to_el = anchor.find(f"{{{xdr}}}to")
        
        from_col = int(from_el.find(f"{{{xdr}}}col").text) if from_el is not None and from_el.find(f"{{{xdr}}}col") is not None else None
        from_row = int(from_el.find(f"{{{xdr}}}row").text) if from_el is not None and from_el.find(f"{{{xdr}}}row") is not None else None
        to_col = int(to_el.find(f"{{{xdr}}}col").text) if to_el is not None and to_el.find(f"{{{xdr}}}col") is not None else None
        to_row = int(to_el.find(f"{{{xdr}}}row").text) if to_el is not None and to_el.find(f"{{{xdr}}}row") is not None else None
        
        # Get size
        ext = anchor.find(f".//{{{a_ns}}}ext")
        width_emu = int(ext.get("cx", 0)) if ext is not None else None
        height_emu = int(ext.get("cy", 0)) if ext is not None else None
        
        # Resolve media path from drawing rels
        drawing_rels_path = drawing_path.replace("drawings/", "drawings/_rels/").replace(".xml", ".xml.rels")
        media_path = ""
        
        try:
            with zf.open(drawing_rels_path) as f:
                drawing_rels_tree = ET.parse(f)
                drawing_rels_root = drawing_rels_tree.getroot()
                
                for rel in drawing_rels_root.findall(f"{{{ns_rel}}}Relationship"):
                    if rel.get("Id") == r_embed:
                        target = rel.get("Target", "")
                        if target.startswith("../"):
                            media_path = "xl/" + target[3:]
                        else:
                            media_path = "xl/drawings/" + target
                        break
        except KeyError:
            pass
        
        images.append(ExcelImage(
            id=f"{sheet_id}-img-{i}",
            name=name,
            description=description,
            anchor_type=anchor_type,
            from_col=from_col,
            from_row=from_row,
            to_col=to_col,
            to_row=to_row,
            width_emu=width_emu,
            height_emu=height_emu,
            media_path=media_path,
            r_id=r_embed,
        ))
    
    return images


def _parse_comments(zf: zipfile.ZipFile, sheet_path: str, sheet_id: str) -> List[ExcelComment]:
    """Parse comments from VML drawings."""
    comments: List[ExcelComment] = []
    
    # Find the VML relationship
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    
    try:
        with zf.open(rels_path) as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
    except KeyError:
        return comments
    
    # Find comments file
    comments_path = None
    ns_rel = NS["rel"]
    for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "comments" in rel_type.lower():
            target = rel.get("Target", "")
            if target.startswith("../"):
                comments_path = "xl/" + target[3:]
            else:
                comments_path = "xl/worksheets/" + target
            break
    
    if not comments_path:
        return comments
    
    try:
        with zf.open(comments_path) as f:
            comments_tree = ET.parse(f)
            comments_root = comments_tree.getroot()
    except KeyError:
        return comments
    
    ns = NS["main"]
    
    # Parse authors
    authors: List[str] = []
    authors_el = comments_root.find(f"{{{ns}}}authors")
    if authors_el is not None:
        for author in authors_el.findall(f"{{{ns}}}author"):
            authors.append(author.text or "")
    
    # Parse comments
    comment_list = comments_root.find(f"{{{ns}}}commentList")
    if comment_list is not None:
        for i, comment in enumerate(comment_list.findall(f"{{{ns}}}comment")):
            cell_ref = comment.get("ref", "")
            author_id = int(comment.get("authorId", 0))
            author = authors[author_id] if author_id < len(authors) else None
            
            # Get text
            text_el = comment.find(f"{{{ns}}}text")
            text_parts = []
            if text_el is not None:
                for r in text_el.findall(f"{{{ns}}}r"):
                    t = r.find(f"{{{ns}}}t")
                    if t is not None and t.text:
                        text_parts.append(t.text)
                # Also check for direct <t>
                for t in text_el.findall(f"{{{ns}}}t"):
                    if t.text:
                        text_parts.append(t.text)
            
            comments.append(ExcelComment(
                id=f"{sheet_id}-comment-{i}",
                cell_ref=cell_ref,
                author=author,
                text="".join(text_parts),
            ))
    
    return comments


# =============================================================================
# HYPERLINKS
# =============================================================================

def _parse_hyperlinks(
    zf: zipfile.ZipFile,
    sheet_el: ET.Element,
    sheet_path: str,
    sheet_id: str,
) -> List[ExcelHyperlink]:
    """Parse hyperlinks from a worksheet."""
    hyperlinks: List[ExcelHyperlink] = []
    ns = NS["main"]
    r_ns = NS["r"]
    
    hl_el = sheet_el.find(f"{{{ns}}}hyperlinks")
    if hl_el is None:
        return hyperlinks
    
    # Load relationship file to resolve external links
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    id_to_target: Dict[str, str] = {}
    
    try:
        with zf.open(rels_path) as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
            ns_rel = NS["rel"]
            for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
                rel_id = rel.get("Id")
                target = rel.get("Target", "")
                if rel_id:
                    id_to_target[rel_id] = target
    except KeyError:
        pass
    
    for i, hl in enumerate(hl_el.findall(f"{{{ns}}}hyperlink")):
        cell_ref = hl.get("ref", "")
        r_id = hl.get(f"{{{r_ns}}}id")
        location = hl.get("location")
        display = hl.get("display")
        tooltip = hl.get("tooltip")
        
        # Resolve target
        target = None
        if r_id:
            target = id_to_target.get(r_id)
        
        hyperlinks.append(ExcelHyperlink(
            id=f"{sheet_id}-hl-{i}",
            cell_ref=cell_ref,
            target=target,
            location=location,
            display=display,
            tooltip=tooltip,
            r_id=r_id,
        ))
    
    return hyperlinks


# =============================================================================
# CONDITIONAL FORMATTING
# =============================================================================

def _parse_conditional_formatting(
    sheet_el: ET.Element,
    sheet_id: str,
) -> List[ConditionalFormatting]:
    """Parse conditional formatting rules from a worksheet."""
    cf_list: List[ConditionalFormatting] = []
    ns = NS["main"]
    
    for i, cf_el in enumerate(sheet_el.findall(f"{{{ns}}}conditionalFormatting")):
        sqref = cf_el.get("sqref", "")
        rules: List[ConditionalFormatRule] = []
        
        for j, rule_el in enumerate(cf_el.findall(f"{{{ns}}}cfRule")):
            rule_type = rule_el.get("type", "")
            priority = int(rule_el.get("priority", 1))
            operator = rule_el.get("operator")
            dxf_id = int(rule_el.get("dxfId")) if rule_el.get("dxfId") else None
            stop_if_true = rule_el.get("stopIfTrue") == "1"
            
            # Formulas
            formula1 = None
            formula2 = None
            formulas = rule_el.findall(f"{{{ns}}}formula")
            if len(formulas) >= 1 and formulas[0].text:
                formula1 = formulas[0].text
            if len(formulas) >= 2 and formulas[1].text:
                formula2 = formulas[1].text
            
            # Color scale
            color_scale = None
            cs_el = rule_el.find(f"{{{ns}}}colorScale")
            if cs_el is not None:
                color_scale = {"cfvos": [], "colors": []}
                for cfvo in cs_el.findall(f"{{{ns}}}cfvo"):
                    color_scale["cfvos"].append({
                        "type": cfvo.get("type"),
                        "val": cfvo.get("val"),
                    })
                for color in cs_el.findall(f"{{{ns}}}color"):
                    color_scale["colors"].append({
                        "rgb": color.get("rgb"),
                        "theme": color.get("theme"),
                        "tint": color.get("tint"),
                    })
            
            # Data bar
            data_bar = None
            db_el = rule_el.find(f"{{{ns}}}dataBar")
            if db_el is not None:
                data_bar = {
                    "min_length": int(db_el.get("minLength", 10)),
                    "max_length": int(db_el.get("maxLength", 90)),
                    "show_value": db_el.get("showValue") != "0",
                }
                color = db_el.find(f"{{{ns}}}color")
                if color is not None:
                    data_bar["color"] = color.get("rgb") or color.get("theme")
            
            # Icon set
            icon_set = None
            is_el = rule_el.find(f"{{{ns}}}iconSet")
            if is_el is not None:
                icon_set = {
                    "icon_set": is_el.get("iconSet", "3TrafficLights1"),
                    "show_value": is_el.get("showValue") != "0",
                    "reverse": is_el.get("reverse") == "1",
                    "cfvos": [],
                }
                for cfvo in is_el.findall(f"{{{ns}}}cfvo"):
                    icon_set["cfvos"].append({
                        "type": cfvo.get("type"),
                        "val": cfvo.get("val"),
                    })
            
            rules.append(ConditionalFormatRule(
                id=f"{sheet_id}-cf-{i}-rule-{j}",
                type=rule_type,
                priority=priority,
                operator=operator,
                formula1=formula1,
                formula2=formula2,
                dxf_id=dxf_id,
                color_scale=color_scale,
                data_bar=data_bar,
                icon_set=icon_set,
                stop_if_true=stop_if_true,
            ))
        
        cf_list.append(ConditionalFormatting(
            id=f"{sheet_id}-cf-{i}",
            sqref=sqref,
            rules=rules,
        ))
    
    return cf_list


# =============================================================================
# FORM CONTROLS (from VML)
# =============================================================================

def _parse_form_controls(
    zf: zipfile.ZipFile,
    sheet_path: str,
    sheet_id: str,
) -> List[FormControl]:
    """Parse form controls from VML drawings."""
    controls: List[FormControl] = []
    
    # Find VML relationship
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    
    try:
        with zf.open(rels_path) as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
    except KeyError:
        return controls
    
    # Find VML file
    vml_path = None
    ns_rel = NS["rel"]
    for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "vmlDrawing" in rel_type.lower():
            target = rel.get("Target", "")
            if target.startswith("../"):
                vml_path = "xl/" + target[3:]
            else:
                vml_path = "xl/worksheets/" + target
            break
    
    if not vml_path:
        return controls
    
    try:
        vml_content = zf.read(vml_path).decode("utf-8", errors="ignore")
    except KeyError:
        return controls
    
    # Parse VML for form controls (basic parsing - VML is messy)
    # Look for x:ClientData with ObjectType
    import re
    
    # Map object types to FormControlType
    type_map = {
        "Checkbox": FormControlType.CHECKBOX,
        "Radio": FormControlType.RADIO,
        "Button": FormControlType.BUTTON,
        "Drop": FormControlType.DROPDOWN,
        "List": FormControlType.LISTBOX,
        "Spin": FormControlType.SPINNER,
        "Scroll": FormControlType.SCROLLBAR,
        "GBox": FormControlType.GROUPBOX,
        "Label": FormControlType.LABEL,
    }
    
    # Find all shapes with form control client data
    shape_pattern = re.compile(
        r'<v:shape[^>]*id="([^"]*)"[^>]*>.*?<x:ClientData\s+ObjectType="([^"]*)">(.*?)</x:ClientData>',
        re.DOTALL | re.IGNORECASE
    )
    
    for i, match in enumerate(shape_pattern.finditer(vml_content)):
        shape_id = match.group(1)
        obj_type = match.group(2)
        client_data = match.group(3)
        
        if obj_type not in type_map:
            continue
        
        ctrl_type = type_map[obj_type]
        
        # Parse client data for properties
        checked = None
        linked_cell = None
        input_range = None
        
        # Checked state (for checkbox/radio)
        if "<x:Checked/>" in client_data:
            checked = True
        elif '<x:Checked>' in client_data:
            checked_match = re.search(r'<x:Checked>(.*?)</x:Checked>', client_data)
            if checked_match:
                checked = checked_match.group(1) != "0"
        
        # Linked cell
        fmla_link = re.search(r'<x:FmlaLink>(.*?)</x:FmlaLink>', client_data)
        if fmla_link:
            linked_cell = fmla_link.group(1)
        
        # Input range (for dropdowns)
        fmla_range = re.search(r'<x:FmlaRange>(.*?)</x:FmlaRange>', client_data)
        if fmla_range:
            input_range = fmla_range.group(1)
        
        # Row/column position
        anchor = re.search(r'<x:Anchor>(.*?)</x:Anchor>', client_data)
        from_col = None
        from_row = None
        to_col = None
        to_row = None
        if anchor:
            parts = [int(x.strip()) for x in anchor.group(1).split(",") if x.strip().isdigit()]
            if len(parts) >= 8:
                from_col = parts[0]
                from_row = parts[2]
                to_col = parts[4]
                to_row = parts[6]
        
        controls.append(FormControl(
            id=f"{sheet_id}-ctrl-{i}",
            name=shape_id,
            control_type=ctrl_type,
            from_col=from_col,
            from_row=from_row,
            to_col=to_col,
            to_row=to_row,
            checked=checked,
            linked_cell=linked_cell,
            input_range=input_range,
            vml_shape_id=shape_id,
        ))
    
    return controls


# =============================================================================
# SHEET VIEW / FREEZE PANES
# =============================================================================

def _parse_sheet_view(
    sheet_el: ET.Element,
    sheet_id: str,
) -> Optional[SheetView]:
    """Parse sheet view settings including freeze panes."""
    ns = NS["main"]
    
    views_el = sheet_el.find(f"{{{ns}}}sheetViews")
    if views_el is None:
        return None
    
    view_el = views_el.find(f"{{{ns}}}sheetView")
    if view_el is None:
        return None
    
    # Basic view settings
    view = SheetView(
        id=f"{sheet_id}-view",
        view_type=view_el.get("view", "normal"),
        zoom_scale=int(view_el.get("zoomScale", 100)),
        zoom_scale_normal=int(view_el.get("zoomScaleNormal")) if view_el.get("zoomScaleNormal") else None,
        zoom_scale_page_layout_view=int(view_el.get("zoomScalePageLayoutView")) if view_el.get("zoomScalePageLayoutView") else None,
        show_gridlines=view_el.get("showGridLines") != "0",
        show_row_col_headers=view_el.get("showRowColHeaders") != "0",
        show_formulas=view_el.get("showFormulas") == "1",
        show_zeros=view_el.get("showZeros") != "0",
    )
    
    # Selection
    selection_el = view_el.find(f"{{{ns}}}selection")
    if selection_el is not None:
        view.active_cell = selection_el.get("activeCell")
        view.active_cell_id = int(selection_el.get("activeCellId", 0)) if selection_el.get("activeCellId") else None
    
    # Freeze pane
    pane_el = view_el.find(f"{{{ns}}}pane")
    if pane_el is not None:
        x_split = int(float(pane_el.get("xSplit", 0)))
        y_split = int(float(pane_el.get("ySplit", 0)))
        
        # Check if it's a freeze (state="frozen") vs split
        state = pane_el.get("state", "")
        if state in ("frozen", "frozenSplit"):
            view.freeze_pane = FreezePane(
                x_split=x_split,
                y_split=y_split,
                top_left_cell=pane_el.get("topLeftCell"),
                active_pane=pane_el.get("activePane", "bottomRight"),
            )
        else:
            # Regular split
            view.split_horizontal = int(float(pane_el.get("ySplit", 0)))
            view.split_vertical = int(float(pane_el.get("xSplit", 0)))
    
    return view


# =============================================================================
# TABLES
# =============================================================================

def _parse_tables(
    zf: zipfile.ZipFile,
    sheet_path: str,
    sheet_id: str,
) -> List[ExcelTable]:
    """Parse structured tables from table XML files."""
    tables: List[ExcelTable] = []
    
    # Find table relationships
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(".xml", ".xml.rels")
    
    try:
        with zf.open(rels_path) as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
    except KeyError:
        return tables
    
    ns_rel = NS["rel"]
    ns = NS["main"]
    
    for rel in rels_root.findall(f"{{{ns_rel}}}Relationship"):
        rel_type = rel.get("Type", "")
        if "table" not in rel_type.lower():
            continue
        
        target = rel.get("Target", "")
        r_id = rel.get("Id")
        
        if target.startswith("../"):
            table_path = "xl/" + target[3:]
        else:
            table_path = "xl/worksheets/" + target
        
        try:
            with zf.open(table_path) as f:
                table_tree = ET.parse(f)
                table_root = table_tree.getroot()
        except KeyError:
            continue
        
        # Parse table
        table_id = table_root.get("id", "")
        name = table_root.get("name", "")
        display_name = table_root.get("displayName", name)
        ref = table_root.get("ref", "")
        header_count = int(table_root.get("headerRowCount", 1))
        totals_count = int(table_root.get("totalsRowCount", 0))
        
        # Parse columns
        columns: List[TableColumn] = []
        cols_el = table_root.find(f"{{{ns}}}tableColumns")
        if cols_el is not None:
            for col_el in cols_el.findall(f"{{{ns}}}tableColumn"):
                columns.append(TableColumn(
                    id=int(col_el.get("id", 0)),
                    name=col_el.get("name", ""),
                    data_cell_style=col_el.get("dataCellStyle"),
                    header_cell_style=col_el.get("headerRowCellStyle"),
                    totals_cell_style=col_el.get("totalsRowCellStyle"),
                    totals_row_function=col_el.get("totalsRowFunction"),
                    totals_row_formula=col_el.get("totalsRowFormula"),
                ))
        
        # Parse style
        style_el = table_root.find(f"{{{ns}}}tableStyleInfo")
        style_name = None
        show_first = False
        show_last = False
        show_row_stripes = True
        show_col_stripes = False
        if style_el is not None:
            style_name = style_el.get("name")
            show_first = style_el.get("showFirstColumn") == "1"
            show_last = style_el.get("showLastColumn") == "1"
            show_row_stripes = style_el.get("showRowStripes") != "0"
            show_col_stripes = style_el.get("showColumnStripes") == "1"
        
        # Auto filter
        af_el = table_root.find(f"{{{ns}}}autoFilter")
        af_ref = af_el.get("ref") if af_el is not None else None
        
        tables.append(ExcelTable(
            id=f"{sheet_id}-table-{table_id}",
            name=name,
            display_name=display_name,
            ref=ref,
            header_row_count=header_count,
            totals_row_count=totals_count,
            columns=columns,
            table_style_name=style_name,
            show_first_column=show_first,
            show_last_column=show_last,
            show_row_stripes=show_row_stripes,
            show_column_stripes=show_col_stripes,
            auto_filter_ref=af_ref,
            r_id=r_id,
        ))
    
    return tables


# =============================================================================
# SPARKLINES
# =============================================================================

def _parse_sparklines(
    sheet_el: ET.Element,
    sheet_id: str,
) -> List[SparklineGroup]:
    """Parse sparklines from extended worksheet elements."""
    groups: List[SparklineGroup] = []
    ns = NS["main"]
    x14_ns = NS["x14"]
    
    # Sparklines are in extLst -> ext -> x14:sparklineGroups
    ext_lst = sheet_el.find(f"{{{ns}}}extLst")
    if ext_lst is None:
        return groups
    
    for ext in ext_lst.findall(f"{{{ns}}}ext"):
        sg_root = ext.find(f"{{{x14_ns}}}sparklineGroups")
        if sg_root is None:
            continue
        
        for i, sg in enumerate(sg_root.findall(f"{{{x14_ns}}}sparklineGroup")):
            sparkline_type = sg.get("type", "line")
            display_empty = sg.get("displayEmptyCellsAs", "gap")
            date_axis = sg.get("dateAxis") == "1"
            
            sparklines: List[Sparkline] = []
            
            sparklines_el = sg.find(f"{{{x14_ns}}}sparklines")
            if sparklines_el is not None:
                for j, sp in enumerate(sparklines_el.findall(f"{{{x14_ns}}}sparkline")):
                    f_el = sp.find(f"{{{x14_ns}}}f")
                    sqref_el = sp.find(f"{{{x14_ns}}}sqref")
                    
                    data_range = f_el.text if f_el is not None and f_el.text else ""
                    location = sqref_el.text if sqref_el is not None and sqref_el.text else ""
                    
                    sparklines.append(Sparkline(
                        id=f"{sheet_id}-spark-{i}-{j}",
                        location=location,
                        data_range=data_range,
                        sparkline_type=sparkline_type,
                    ))
            
            groups.append(SparklineGroup(
                id=f"{sheet_id}-sparkgroup-{i}",
                sparklines=sparklines,
                sparkline_type=sparkline_type,
                display_empty_cells_as=display_empty,
                date_axis=date_axis,
            ))
    
    return groups


# =============================================================================
# DEFINED NAMES
# =============================================================================

def _parse_defined_names(wb_root: ET.Element) -> List[DefinedName]:
    """Parse defined names from workbook.xml."""
    names: List[DefinedName] = []
    ns = NS["main"]
    
    dn_el = wb_root.find(f"{{{ns}}}definedNames")
    if dn_el is None:
        return names
    
    for i, dn in enumerate(dn_el.findall(f"{{{ns}}}definedName")):
        name = dn.get("name", "")
        value = dn.text or ""
        local_sheet = dn.get("localSheetId")
        hidden = dn.get("hidden") == "1"
        comment = dn.get("comment")
        
        # Check if built-in name
        is_builtin = name.startswith("_xlnm.")
        
        names.append(DefinedName(
            id=f"dn-{i}",
            name=name,
            value=value,
            local_sheet_id=int(local_sheet) if local_sheet else None,
            hidden=hidden,
            comment=comment,
            is_builtin=is_builtin,
        ))
    
    return names


def parse_sheet(
    zf: zipfile.ZipFile,
    sheet_path: str,
    sheet_name: str,
    sheet_index: int,
    shared_strings: List[SharedStringItem],
    style_info: StyleInfo,
    is_hidden: bool = False,
) -> ExcelSheetJSON:
    """Parse a single worksheet from the XLSX archive."""
    
    sheet_id = f"sheet-{sheet_index}"
    
    with zf.open(sheet_path) as f:
        tree = ET.parse(f)
        sheet_el = tree.getroot()
    
    ns = NS["main"]
    
    # Get dimension
    dimension_el = sheet_el.find(f"{{{ns}}}dimension")
    dimension = dimension_el.get("ref") if dimension_el is not None else None
    
    # Parse merged cells first (needed for cell parsing)
    merged_cells = _parse_merged_cells(sheet_el, sheet_id)
    
    # Parse cells and rows
    cells, rows = _parse_cells(sheet_el, sheet_id, shared_strings, style_info, merged_cells)
    
    # Parse other elements
    columns = _parse_columns(sheet_el)
    data_validations = _parse_data_validations(sheet_el, sheet_id)
    images = _parse_images(zf, sheet_path, sheet_id)
    comments = _parse_comments(zf, sheet_path, sheet_id)
    
    # Parse new complex elements
    hyperlinks = _parse_hyperlinks(zf, sheet_el, sheet_path, sheet_id)
    conditional_formatting = _parse_conditional_formatting(sheet_el, sheet_id)
    form_controls = _parse_form_controls(zf, sheet_path, sheet_id)
    tables = _parse_tables(zf, sheet_path, sheet_id)
    sparkline_groups = _parse_sparklines(sheet_el, sheet_id)
    sheet_view = _parse_sheet_view(sheet_el, sheet_id)
    
    return ExcelSheetJSON(
        id=sheet_id,
        name=sheet_name,
        sheet_index=sheet_index,
        is_hidden=is_hidden,
        dimension=dimension,
        cells=cells,
        merged_cells=merged_cells,
        data_validations=data_validations,
        columns=columns,
        rows=rows,
        images=images,
        comments=comments,
        # New complex elements
        hyperlinks=hyperlinks,
        conditional_formatting=conditional_formatting,
        form_controls=form_controls,
        tables=tables,
        sparkline_groups=sparkline_groups,
        sheet_view=sheet_view,
    )


# =============================================================================
# WORKBOOK PARSING
# =============================================================================

def xlsx_to_json(xlsx_path: str, workbook_id: str) -> ExcelWorkbookJSON:
    """Parse an XLSX file into ExcelWorkbookJSON structure.
    
    Handles:
    - Multiple worksheets
    - Merged cells
    - Data validation (dropdowns)
    - Cell formatting and styles
    - Images
    - Comments
    - Formulas
    """
    
    xlsx_path = str(xlsx_path)
    
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        # Parse shared strings
        shared_strings = _parse_shared_strings(zf)
        
        # Parse styles
        style_info = _parse_styles(zf)
        
        # Parse workbook.xml for sheet info
        with zf.open("xl/workbook.xml") as f:
            wb_tree = ET.parse(f)
            wb_root = wb_tree.getroot()
        
        ns = NS["main"]
        r_ns = NS["r"]
        
        # Get sheet names and r:ids
        sheets_el = wb_root.find(f"{{{ns}}}sheets")
        sheet_infos: List[Dict[str, Any]] = []
        
        if sheets_el is not None:
            for sheet in sheets_el.findall(f"{{{ns}}}sheet"):
                sheet_infos.append({
                    "name": sheet.get("name"),
                    "sheetId": sheet.get("sheetId"),
                    "r_id": sheet.get(f"{{{r_ns}}}id"),
                    "state": sheet.get("state"),
                })
        
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
        
        # Parse each sheet
        sheets: List[ExcelSheetJSON] = []
        for i, info in enumerate(sheet_infos):
            r_id = info.get("r_id")
            target = id_to_target.get(r_id, "")
            
            # Build full path
            if target.startswith("/"):
                sheet_path = target[1:]
            else:
                sheet_path = f"xl/{target}"
            
            is_hidden = info.get("state") == "hidden"
            
            try:
                sheet = parse_sheet(
                    zf,
                    sheet_path,
                    info["name"],
                    i,
                    shared_strings,
                    style_info,
                    is_hidden,
                )
                sheets.append(sheet)
            except KeyError as e:
                # Sheet file not found
                continue
        
        # Get active sheet
        book_views = wb_root.find(f"{{{ns}}}bookViews")
        active_sheet = 0
        if book_views is not None:
            wv = book_views.find(f"{{{ns}}}workbookView")
            if wv is not None:
                active_sheet = int(wv.get("activeTab", 0))
        
        # Parse defined names
        defined_names = _parse_defined_names(wb_root)
        
        # Parse core properties for metadata
        created = None
        modified = None
        creator = None
        last_modified_by = None
        
        try:
            with zf.open("docProps/core.xml") as f:
                core_tree = ET.parse(f)
                core_root = core_tree.getroot()
                
                dc = NS["dc"]
                dcterms = NS["dcterms"]
                cp = NS["cp"]
                
                creator_el = core_root.find(f"{{{dc}}}creator")
                if creator_el is not None:
                    creator = creator_el.text
                
                created_el = core_root.find(f"{{{dcterms}}}created")
                if created_el is not None:
                    created = created_el.text
                
                modified_el = core_root.find(f"{{{dcterms}}}modified")
                if modified_el is not None:
                    modified = modified_el.text
                
                last_mod_el = core_root.find(f"{{{cp}}}lastModifiedBy")
                if last_mod_el is not None:
                    last_modified_by = last_mod_el.text
        except KeyError:
            pass
    
    return ExcelWorkbookJSON(
        id=workbook_id,
        filename=Path(xlsx_path).name,
        sheets=sheets,
        active_sheet_index=active_sheet,
        shared_strings=shared_strings,
        defined_names=defined_names,
        created=created,
        modified=modified,
        creator=creator,
        last_modified_by=last_modified_by,
    )
