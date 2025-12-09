# API Commands for DOCX Pipeline Testing

This document provides curl commands for testing the document pipeline with `test2.docx`.

> **Note**: On Windows PowerShell, use `curl.exe` instead of `curl` (PowerShell aliases `curl` to `Invoke-WebRequest`).

---

## Prerequisites

1. **Start the server**:
   ```powershell
   .venv\Scripts\Activate.ps1
   uvicorn main:app --reload --port 8000
   ```

2. **Test file location**: `data/uploads/docx/test2.docx`

---

## Pipeline Overview

```
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│   DOCX      │───>│   XML       │───>│   JSON      │───>│   DOCX      │
│  (upload)   │    │  (internal) │    │  (editable) │    │  (export)   │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
      │                  │                  │                  │
   Step 1             Step 2             Step 3             Step 4
   Upload           Parse XML          Edit JSON          Merge back
```

---

## Step 1: Upload DOCX → Parse to JSON

**Endpoint**: `POST /documents/`

```bash
# Upload test2.docx and receive JSON structure
curl.exe -X POST "http://127.0.0.1:8000/documents/" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@data/uploads/docx/test2.docx"
```

**PowerShell alternative**:
```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:8000/documents/" `
  -Method POST `
  -Form @{ file = Get-Item "data\uploads\docx\test2.docx" }
```

**Response**: Returns `DocumentJSON` with:
- `id`: Document identifier (filename)
- `blocks`: Array of paragraphs, tables, and drawings
- `checkboxes`: Array of checkbox form fields
- `dropdowns`: Array of dropdown form fields

---

## Step 2: Get Document JSON

**Endpoint**: `GET /documents/{document_id}`

```bash
# Retrieve the parsed JSON structure
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx"
```

**Response**: Full `DocumentJSON` object.

---

## Step 3: Update Document JSON (Edit)

**Endpoint**: `PUT /documents/{document_id}`

```bash
# First, get the current JSON and save it
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx" -o document.json

# Edit document.json in your editor, then update:
curl.exe -X PUT "http://127.0.0.1:8000/documents/test2.docx" \
  -H "Content-Type: application/json" \
  -d @document.json
```

**Example: Inline edit (changing text in first paragraph)**:
```bash
# Note: This requires a valid JSON body matching DocumentJSON schema
# The full JSON structure must be provided
```

---

## Step 4: Export JSON → DOCX

**Endpoint**: `POST /documents/{document_id}/export`

```bash
# Export and get metadata response
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/export"
```

**Response**: Returns JSON with `export_path` and `version`.

---

## Step 5: Download Exported DOCX File

**Endpoint**: `POST /documents/{document_id}/export/file`

```bash
# Download the exported DOCX file directly
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/export/file" \
  --output test2_exported.docx
```

This downloads the reconstructed DOCX file to your local directory.

---

## Validation Endpoints

### Validate Current Document State

**Endpoint**: `GET /documents/{document_id}/validate`

```bash
# Validate the document against original DOCX
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx/validate"
```

**Response**: Validation report with:
- `has_errors`: Boolean
- `has_warnings`: Boolean
- `stages`: Array of stage snapshots (char counts, paragraph counts, etc.)
- `issues`: Array of validation issues

### Validate Export (Dry Run)

**Endpoint**: `POST /documents/{document_id}/validate-export`

```bash
# Test export without saving
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/validate-export"
```

---

## Form Field Endpoints

### Update Checkbox

**Endpoint**: `POST /documents/{document_id}/checkbox`

```bash
# Toggle a checkbox (get checkbox IDs from document JSON)
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/checkbox" \
  -H "Content-Type: application/json" \
  -d "{\"checkbox_id\": \"checkbox-12345\", \"checked\": true}"
```

### Update Dropdown

**Endpoint**: `POST /documents/{document_id}/dropdown`

```bash
# Select a dropdown option
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/dropdown" \
  -H "Content-Type: application/json" \
  -d "{\"dropdown_id\": \"dropdown-67890\", \"selected\": \"Yes\"}"
```

---

## AI-Powered Edit

**Endpoint**: `POST /documents/{document_id}/ai-edit`

```bash
# Apply AI edit to a block (requires AI provider configured)
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/ai-edit" \
  -H "Content-Type: application/json" \
  -d "{\"block_id\": \"p-0\", \"instruction\": \"make more formal\"}"
```

**For table cell edits**:
```bash
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/ai-edit" \
  -H "Content-Type: application/json" \
  -d "{\"block_id\": \"tbl-0\", \"cell_id\": \"cell-0-0-0\", \"instruction\": \"uppercase\"}"
```

---

## HTML Preview

**Endpoint**: `GET /documents/{document_id}/preview/html`

```bash
# Get HTML preview (opens in browser)
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx/preview/html"
```

Or simply open in browser: `http://127.0.0.1:8000/documents/test2.docx/preview/html`

---

## Complete Workflow Example

```bash
# 1. Upload document
curl.exe -X POST "http://127.0.0.1:8000/documents/" \
  -F "file=@data/uploads/docx/test2.docx"

# 2. View document structure
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx"

# 3. Validate parsing
curl.exe -X GET "http://127.0.0.1:8000/documents/test2.docx/validate"

# 4. Export back to DOCX
curl.exe -X POST "http://127.0.0.1:8000/documents/test2.docx/export/file" \
  --output test2_roundtrip.docx

# 5. Open test2_roundtrip.docx in Microsoft Word to verify
```

---

## Programmatic Access (Python)

```python
import requests

BASE_URL = "http://127.0.0.1:8000"

# Upload
with open("data/uploads/docx/test2.docx", "rb") as f:
    response = requests.post(f"{BASE_URL}/documents/", files={"file": f})
doc_json = response.json()

# Get document
response = requests.get(f"{BASE_URL}/documents/test2.docx")
doc = response.json()

# Modify a run
doc["blocks"][0]["runs"][0]["text"] += " [EDITED]"

# Update
response = requests.put(
    f"{BASE_URL}/documents/test2.docx",
    json=doc
)

# Export
response = requests.post(
    f"{BASE_URL}/documents/test2.docx/export/file"
)
with open("test2_exported.docx", "wb") as f:
    f.write(response.content)
```

---

## Debugging Intermediate Results

To see intermediate XML and JSON at each pipeline stage, use the debug scripts:

```bash
# Full pipeline debug
python debug/docx/debug_pipeline.py

# Comprehensive fidelity test
python debug/comprehensive_fidelity_test.py

# Backend E2E test
python tests/test_backend_e2e.py data/uploads/docx/test2.docx
```

---

## Response Codes

| Code | Meaning |
|------|---------|
| 200 | Success |
| 400 | Validation error / Bad request |
| 404 | Document not found |
| 500 | Server error |

---

## Notes

1. **Document ID**: The document ID is the filename (e.g., `test2.docx`)
2. **Export versioning**: Each export increments the version number
3. **XML preservation**: The pipeline preserves all DOCX XML structure except `word/document.xml` which is patched
4. **Form fields**: Checkboxes and dropdowns are extracted separately from block content
