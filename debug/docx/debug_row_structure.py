"""Check row structure for SDT wrapping."""

import zipfile
from xml.etree import ElementTree as ET

with zipfile.ZipFile('data/uploads/docx/test2.docx') as zf:
    tree = ET.parse(zf.open('word/document.xml'))

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
root = tree.getroot()
body = root.find('w:body', NS)

# Find tbl[1]
tables = [c for c in body if c.tag.endswith('}tbl')]
print(f'Found {len(tables)} tables')

tbl = tables[1]  # tbl[1]
rows = [c for c in tbl if c.tag.endswith('}tr')]
print(f'Table has {len(rows)} direct tr children')

# Check row 4
row = rows[4]
print(f'\nRow 4 direct children:')
for i, child in enumerate(row):
    tag = child.tag.split('}')[1]
    print(f'  [{i}] {tag}')
    
# Check if there are tc inside sdt
sdts = row.findall('.//w:sdt', NS)
print(f'\nRow has {len(sdts)} SDT elements')
tcs = row.findall('.//w:tc', NS)
print(f'Row has {len(tcs)} tc elements (including nested)')

# Show the structure
print('\nDetailed structure of row 4:')
def show_structure(el, indent=0):
    tag = el.tag.split('}')[1]
    text = ''
    if tag == 't' and el.text:
        text = f" = '{el.text[:30]}'"
    print('  ' * indent + tag + text)
    for child in el:
        show_structure(child, indent + 1)

# Just show first few levels
for i, child in enumerate(row):
    tag = child.tag.split('}')[1]
    print(f'\n[{i}] {tag}:')
    for j, grandchild in enumerate(child):
        gtag = grandchild.tag.split('}')[1]
        print(f'    [{j}] {gtag}')
        if gtag == 'sdtContent':
            for k, ggchild in enumerate(grandchild):
                ggtag = ggchild.tag.split('}')[1]
                print(f'        [{k}] {ggtag}')
