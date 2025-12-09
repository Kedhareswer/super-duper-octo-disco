"""Analyze merged cells, borders, and drawings in DOCX."""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
}

with zipfile.ZipFile('test2.docx', 'r') as z:
    with z.open("word/document.xml") as doc_xml:
        tree = ET.parse(doc_xml)

root = tree.getroot()
body = root.find("w:body", NS)

# Analyze drawings first
drawings = body.findall(".//w:drawing", NS)
print(f"=== ANALYZING {len(drawings)} DRAWINGS ===\n")

for di, drawing in enumerate(drawings):
    print(f"DRAWING {di}")
    print("-" * 40)
    
    # Check for inline or anchor
    inline = drawing.find("wp:inline", NS)
    anchor = drawing.find("wp:anchor", NS)
    
    container = inline if inline is not None else anchor
    if container is not None:
        # Get extent (size)
        extent = container.find("wp:extent", NS)
        if extent is not None:
            cx = int(extent.attrib.get("cx", 0)) / 914400  # EMUs to inches
            cy = int(extent.attrib.get("cy", 0)) / 914400
            print(f"  Size: {cx:.2f}\" x {cy:.2f}\"")
        
        # Get docPr (name)
        doc_pr = container.find("wp:docPr", NS)
        if doc_pr is not None:
            name = doc_pr.attrib.get("name", "")
            print(f"  Name: {name}")
        
        # Check if it's a group (vector shapes)
        wgp = container.find(".//wpg:wgp", NS)
        if wgp is not None:
            print("  Type: Vector Group (wpg:wgp)")
            # Count shapes in group
            shapes = wgp.findall(".//*")
            print(f"  Elements: {len(shapes)}")
        else:
            print("  Type: Other (image/chart)")
    print()

tables = body.findall(".//w:tbl", NS)
print(f"=== ANALYZING {len(tables)} TABLES ===\n")

for ti, tbl in enumerate(tables):
    print(f"TABLE {ti}")
    print("-" * 40)
    
    # Check table borders
    tbl_pr = tbl.find("w:tblPr", NS)
    if tbl_pr is not None:
        tbl_borders = tbl_pr.find("w:tblBorders", NS)
        if tbl_borders is not None:
            print("  Table borders:")
            for border in tbl_borders:
                tag = border.tag.split('}')[1]
                val = border.attrib.get(f"{{{NS['w']}}}val", "none")
                sz = border.attrib.get(f"{{{NS['w']}}}sz", "0")
                color = border.attrib.get(f"{{{NS['w']}}}color", "auto")
                print(f"    {tag}: val={val}, sz={sz}, color={color}")
    
    rows = tbl.findall(".//w:tr", NS)
    print(f"  Rows: {len(rows)}")
    
    # Check for merged cells
    merge_info = []
    for ri, row in enumerate(rows):
        cells = row.findall(".//w:tc", NS)
        for ci, cell in enumerate(cells):
            tc_pr = cell.find("w:tcPr", NS)
            if tc_pr is not None:
                # Horizontal merge (gridSpan)
                grid_span = tc_pr.find("w:gridSpan", NS)
                if grid_span is not None:
                    span = grid_span.attrib.get(f"{{{NS['w']}}}val", "1")
                    if int(span) > 1:
                        merge_info.append(f"    Row {ri}, Cell {ci}: colspan={span}")
                
                # Vertical merge
                v_merge = tc_pr.find("w:vMerge", NS)
                if v_merge is not None:
                    val = v_merge.attrib.get(f"{{{NS['w']}}}val", "continue")
                    merge_info.append(f"    Row {ri}, Cell {ci}: vMerge={val}")
                
                # Cell borders
                tc_borders = tc_pr.find("w:tcBorders", NS)
                if tc_borders is not None and ri == 0 and ci == 0:
                    print(f"  Cell borders (sample from R0C0):")
                    for border in tc_borders:
                        tag = border.tag.split('}')[1]
                        val = border.attrib.get(f"{{{NS['w']}}}val", "none")
                        print(f"    {tag}: {val}")
    
    if merge_info:
        print("  Merged cells:")
        for info in merge_info[:10]:  # Limit output
            print(info)
        if len(merge_info) > 10:
            print(f"    ... and {len(merge_info) - 10} more")
    
    print()
