# Document Digital Copy POC – Quickstart Guide

Get the application (DOCX + Excel) running in 5 minutes.

---

## Prerequisites

| Requirement | Version | Check Command |
|-------------|---------|---------------|
| Python | 3.10+ | `python --version` |
| Node.js | 18+ | `node --version` |
| npm | 8+ | `npm --version` |

---

## Step 1: Setup Backend

```powershell
# Navigate to project
cd path\to\poc-2

# Create virtual environment
python -m venv .venv

# Activate (Windows PowerShell)
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

---

## Step 2: Setup Frontend

```powershell
# From project root
cd web
npm install
cd ..
```

---

## Step 3: Configure Environment

Create `.env` file in project root:

```env
# Required for AI features (get from Google AI Studio)
GOOGLE_API_KEY=your-gemini-api-key

# Optional settings
GEMINI_MODEL=gemini-2.5-flash
DISABLE_RATE_LIMIT=1
DISABLE_EVALS=0              # Set to 1 to disable baked-in evals
EVALS_LOG_TO_CONSOLE=1       # Log eval scores to console
```

> **Note:** AI features work without the API key using rule-based fallbacks.
> Evals are now baked into the main edit flow and run automatically on every AI edit.

---

## Step 4: Start the Application

### Terminal 1: Backend

```powershell
cd path\to\poc-2
.venv\Scripts\Activate.ps1
uvicorn main:app --reload --port 8000
```

You should see:
```
INFO:     Uvicorn running on http://127.0.0.1:8000
INFO:     Application startup complete.
```

### Terminal 2: Frontend

```powershell
cd path\to\poc-2\web
npm run dev
```

You should see:
```
▲ Next.js 16.0.5
- Local: http://localhost:3000
```

---

## Step 5: Use the Application

Open http://localhost:3000 in your browser.

### Upload a Document

1. Click **Choose DOCX File** or **Choose Excel File**
2. Select a sample file, for example:
   - DOCX: `data/uploads/docx/test2.docx`
   - Excel: `data/uploads/excel/excel_test.XLSX`
3. Document is parsed and displayed

### Navigate the Interface

```
┌─────────────────────────────────────────────────────────────────┐
│  ┌─────────────┐  ┌─────────────────┐  ┌─────────────────────┐  │
│  │   BLOCKS    │  │     EDITOR      │  │      PREVIEW        │  │
│  │             │  │                 │  │                     │  │
│  │ P0 Para...  │  │ Selected text   │  │ Formatted document  │  │
│  │ T1 Table    │  │ [Edit area]     │  │ with styling        │  │
│  │  R0C0 Cell  │  │                 │  │                     │  │
│  │  R0C1 Cell  │  │ AI Instruction  │  │ Click elements to   │  │
│  │ D2 Drawing  │  │ [Run AI Edit]   │  │ select them         │  │
│  │             │  │                 │  │                     │  │
│  └─────────────┘  └─────────────────┘  └─────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

### Bidirectional Selection

- **Click block in left panel** → Highlights in preview, scrolls to it
- **Click element in preview** → Highlights in blocks panel, scrolls to it

### Edit Content

1. Select a paragraph or table cell
2. Edit text in the Editor panel
3. Click **Save JSON** to persist

### AI-Powered Editing

1. Select a block/cell
2. Enter instruction: `"make more formal"`, `"fix grammar"`, `"summarize"`
3. Click **Run AI Edit**

### Form Controls

- **Checkboxes:** Toggle in the Editor panel
- **Dropdowns:** Select from options in the Editor panel

### Export Document

1. Click **Export DOCX**
2. File downloads with all edits applied
3. Open in Microsoft Word to verify

---

## Verification

### Test the Backend

```powershell
# Run E2E tests
python -m tests.test_backend_e2e

# Test with specific file
python -m tests.test_backend_e2e "data/uploads/test.docx"
```

### Example Output

The exact output will depend on the document you test. For the bundled sample
document, a successful run looks similar to:

```
======================================================================
  BACKEND END-TO-END TEST SUITE
======================================================================
  ✓ Test file exists
  ✓ Parsed successfully
  ✓ Document passes validation
  ✓ Exported successfully
  ✓ Block count preserved
  ✓ Text segments preserved
======================================================================
  TEST SUITE COMPLETE
======================================================================
```

---

## Troubleshooting

### Backend won't start

```powershell
# Check Python version
python --version  # Should be 3.10+

# Reinstall dependencies
pip install -r requirements.txt --force-reinstall
```

### Frontend won't start

```powershell
# Clear node_modules and reinstall
cd web
Remove-Item -Recurse -Force node_modules
npm install
```

### CORS errors in browser

Ensure backend is running on port 8000:
```powershell
uvicorn main:app --reload --port 8000
```

### Rate limit errors

Set environment variable before starting:
```powershell
$env:DISABLE_RATE_LIMIT = "1"
uvicorn main:app --reload --port 8000
```

### AI edit not working

1. Check `GOOGLE_API_KEY` in `.env`
2. Verify key has Gemini API access
3. Check backend console for errors

---

## API Quick Reference

| Action | Endpoint | Method |
|--------|----------|--------|
| Upload DOCX | `/documents/` | POST |
| Get DOCX | `/documents/{id}` | GET |
| Update DOCX | `/documents/{id}` | PUT |
| Export DOCX | `/documents/{id}/export/file` | POST |
| AI Edit DOCX | `/documents/{id}/ai-edit` | POST |
| Upload Excel | `/spreadsheets/` | POST |
| Get Excel | `/spreadsheets/{id}` | GET |
| Update Excel | `/spreadsheets/{id}` | PUT |
| Export Excel | `/spreadsheets/{id}/export/file` | POST |

### Example: Export via API

```powershell
$id = "test.docx"
Invoke-WebRequest `
  -Uri "http://127.0.0.1:8000/documents/$id/export/file" `
  -Method POST `
  -OutFile "exported.docx"
```

---

## Next Steps

1. **Read the docs:** See `docs/` folder for detailed documentation
2. **Run tests:** `pytest tests/ -v`
3. **Explore the code:** Start with `services/document_engine.py`
4. **Understand the AI flow:** See `services/document_edit_service.py` for orchestration
5. **Try the eval dashboard:** http://localhost:3000/evals

---

## File Locations

| Item | Path |
|------|------|
| Uploaded DOCX files | `data/uploads/docx/` |
| Uploaded Excel files | `data/uploads/excel/` |
| Exported files (DOCX + Excel) | `data/exports/` |
| Database | `data/app.db` |
| Test outputs | `data/test_outputs/` |
| Sample DOCX files | `data/uploads/docx/test.docx`, `data/uploads/docx/test2.docx` |
| Sample Excel file | `data/uploads/excel/excel_test.XLSX` |
