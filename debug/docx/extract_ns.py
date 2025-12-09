import zipfile
import re
zf = zipfile.ZipFile('data/uploads/docx/test2.docx')
content = zf.read('word/document.xml').decode('utf-8')
# Find all xmlns declarations
ns_pattern = r'xmlns:(\w+)="([^"]+)"'
matches = re.findall(ns_pattern, content[:3000])
print('Namespaces in original document:')
for prefix, uri in matches:
    print(f'    "{prefix}": "{uri}",')
