# Development Guide

> Complete guide for setting up, developing, testing, and debugging the DiligenceVault Document Processing System

This guide covers everything you need to start developing, from initial setup to advanced debugging techniques.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Project Setup](#project-setup)
3. [Running the Application](#running-the-application)
4. [Project Structure](#project-structure)
5. [Development Workflow](#development-workflow)
6. [Testing](#testing)
7. [Debugging](#debugging)
8. [Common Issues](#common-issues)
9. [Code Style](#code-style)

---

## Prerequisites

### Required Software

| Software | Version | Purpose | Installation |
|----------|---------|---------|--------------|
| **Python** | 3.10+ | Backend runtime | [python.org](https://python.org) |
| **Node.js** | 18+ | Frontend runtime | [nodejs.org](https://nodejs.org) |
| **Git** | Any | Version control | [git-scm.com](https://git-scm.com) |

### Optional Software

| Software | Purpose |
|----------|---------|
| **VS Code** | Recommended IDE with Python/TypeScript extensions |
| **Postman** | API testing |
| **Microsoft Excel** | Testing exported files |
| **Microsoft Word** | Testing exported files |

### API Keys

| Key | Required For | How to Get |
|-----|--------------|------------|
| `GOOGLE_API_KEY` | AI editing features | [Google AI Studio](https://aistudio.google.com/) |

---

## Project Setup

### 1. Clone and Setup Virtual Environment

```powershell
# Clone the repository
git clone <repository-url>
cd poc-2

# Create virtual environment
python -m venv .venv

# Activate (Windows PowerShell)
.venv\Scripts\Activate.ps1

# Activate (Linux/Mac)
source .venv/bin/activate
```

### 2. Install Dependencies

```powershell
# Backend dependencies
pip install -r requirements.txt

# Frontend dependencies
cd web
npm install
cd ..
```

### 3. Configure Environment

Create `.env` file in project root:

```env
# Required for AI features
GOOGLE_API_KEY=your-gemini-api-key

# Optional
GEMINI_MODEL=gemini-2.5-flash
DISABLE_RATE_LIMIT=1  # For development
```

---

## Running the Application

### Start Backend

```powershell
# From project root
.venv\Scripts\Activate.ps1
uvicorn main:app --reload --port 8000
```

Backend will be available at: `http://127.0.0.1:8000`

### Start Frontend

```powershell
# In a new terminal
cd web
npm run dev
```

Frontend will be available at: `http://localhost:3000`

---

## Project Structure

```
poc-2/
├── main.py                 # FastAPI entry point
├── requirements.txt        # Python dependencies
├── .env                    # Environment variables
│
├── api/                    # API routes
│   └── routes/
│       ├── documents.py    # DOCX endpoints
│       ├── spreadsheets.py # Excel endpoints
│       └── evals.py        # Evaluation endpoints
│
├── models/                 # Pydantic schemas
│   └── schemas.py          # DOCX schemas
│
├── services/               # Business logic
│   ├── document_engine.py  # DOCX processing
│   ├── excel_engine/       # Excel processing (isolated)
│   │   ├── schemas.py      # Excel schemas (20+ models)
│   │   ├── parser.py       # XLSX → JSON
│   │   └── writer.py       # JSON → XLSX
│   ├── ai_agent.py         # AI integration
│   └── db.py               # Database layer
│
├── middleware/             # HTTP middleware
│   └── rate_limit.py
│
├── tests/                  # Test suite
│   ├── test_backend_e2e.py
│   ├── test_fidelity.py    # DOCX fidelity
│   ├── test_excel_engine.py# Excel fidelity
│   └── test_export_roundtrip.py
│
├── debug/                  # Debug utilities
│
├── data/                   # Runtime data
│   ├── uploads/            # DOCX/XLSX uploads
│   ├── exports/            # DOCX/XLSX exports
│   └── app.db
│
├── web/                    # Next.js frontend
│   ├── src/app/
│   │   ├── page.tsx        # Main editor (DOCX/Excel toggle)
│   │   └── evals/page.tsx  # Eval dashboard
│   └── package.json
│
└── docs/                   # Documentation
```

---

## Development Workflow

### Adding a New API Endpoint

1. **Define the route** in `api/routes/documents.py` or create a new router:

```python
@router.post("/{document_id}/new-feature")
async def new_feature(document_id: str, payload: NewFeatureRequest):
    # Implementation
    pass
```

2. **Add request/response models** in `models/schemas.py`:

```python
class NewFeatureRequest(BaseModel):
    field1: str
    field2: int
```

3. **Register the router** in `main.py` (if new file):

```python
from api.routes.new_router import router as new_router
app.include_router(new_router)
```

### Adding a New Document Feature

1. **Update the schema** in `models/schemas.py`
2. **Update parsing** in `services/document_engine.py`:
   - Add extraction in `docx_to_json()`
   - Add patching in `apply_json_to_docx()`
3. **Update frontend** in `web/src/app/page.tsx`
4. **Add tests** in `tests/`

### Modifying the AI Agent

1. **Edit state** in `services/ai_agent.py`:

```python
class EditState(TypedDict):
    # Add new fields
    new_field: str
```

2. **Add/modify nodes**:

```python
def new_node(state: EditState) -> EditState:
    # Process state
    return {...state, "new_field": value}
```

3. **Update graph edges**:

```python
workflow.add_node("new_node", new_node)
workflow.add_edge("previous_node", "new_node")
```

---

## Testing

### Run All Tests

```powershell
pytest tests/ -v
```

### Run DOCX Tests

```powershell
pytest tests/test_fidelity.py -v
pytest tests/test_backend_e2e.py -v
```

### Run Excel Tests

```powershell
# Basic Excel tests
python tests/test_excel_engine.py

# Excel fidelity tests (roundtrip verification)
python tests/test_excel_fidelity.py
```

**Excel Fidelity Test Results (Dec 2025):**
```
Test file: 13 sheets, 576,726 cells, 9,010 merges, 117 CF rules, 7,053 comments

✅ Parse Completeness: PASS
✅ Roundtrip No Changes: PASS  
✅ Roundtrip With Edit: PASS
✅ High Fidelity Elements: PASS
```

### Run E2E Test with Custom File

```powershell
python -m tests.test_backend_e2e "data/uploads/your-file.docx"
```

### Test Coverage

```powershell
pytest tests/ --cov=services --cov-report=html
```

---

## Code Style

### Python

- Follow PEP 8
- Use type hints
- Document functions with docstrings

```python
def function_name(param: str) -> ReturnType:
    """Brief description.
    
    Args:
        param: Parameter description.
        
    Returns:
        Return value description.
    """
    pass
```

### TypeScript

- Use strict mode
- Define types for all props and state
- Use functional components with hooks

```typescript
type Props = {
  value: string;
  onChange: (value: string) => void;
};

export function Component({ value, onChange }: Props) {
  // ...
}
```

---

## Debugging

### Backend Debugging

1. **Enable debug logging**:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

2. **Use the debug scripts** in `debug/`:

```powershell
python debug/debug_pipeline.py
```

3. **Check API responses**:

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:8000/documents/test.docx" | ConvertTo-Json -Depth 10
```

### Frontend Debugging

1. **Use React DevTools**
2. **Check browser console** for API errors
3. **Enable verbose logging**:

```typescript
console.log("State:", doc);
```

### DOCX Debugging

1. **Inspect XML directly**:

```powershell
# Extract DOCX
Expand-Archive -Path "test.docx" -DestinationPath "test_extracted"

# View document.xml
Get-Content "test_extracted/word/document.xml"
```

2. **Use debug scripts**:

```powershell
python debug/debug_text_loss.py
```

---

## Common Issues

### Issue: WinError 1450 - Insufficient system resources

**Cause:** Uvicorn's file watcher runs out of file handles watching `.venv` folder.

**Solution:** Run uvicorn without watching `.venv`:

```powershell
# Option 1: Exclude .venv from watch
uvicorn main:app --reload --port 8000 --reload-exclude ".venv"

# Option 2: Run without reload (production-like)
uvicorn main:app --port 8000
```

### Issue: "Module not found" errors

**Solution:** Ensure virtual environment is activated and dependencies installed:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Issue: CORS errors in browser

**Solution:** Ensure backend is running on port 8000:

```powershell
uvicorn main:app --reload --port 8000
```

### Issue: "Unreadable content" in exported DOCX

**Solution:** This was fixed by:
1. Registering all OOXML namespaces with ElementTree
2. Adding `standalone="yes"` to XML declaration

If it recurs, check `services/document_engine.py` namespace handling.

### Issue: Rate limit errors during development

**Solution:** Set environment variable:

```powershell
$env:DISABLE_RATE_LIMIT = "1"
uvicorn main:app --reload --port 8000
```

### Issue: AI edit returns original text

**Solution:** Check:
1. `GOOGLE_API_KEY` is set correctly
2. API key has Gemini access
3. Check console for error messages

---

## Performance Considerations

### Large Documents

For documents with 1000+ blocks:
- Parsing may take 2-5 seconds
- Export may take 3-10 seconds
- Consider pagination in frontend

### Rate Limiting

Default limits:
- 60 requests/minute (general)
- 20 requests/minute (AI endpoints)

Adjust in `main.py`:

```python
rate_config = RateLimitConfig(
    requests_per_minute=120,
    ai_requests_per_minute=30,
)
```

---

## Deployment Considerations

### Production Checklist

- [ ] Set `DISABLE_RATE_LIMIT` to false
- [ ] Restrict CORS origins
- [ ] Use production database (PostgreSQL)
- [ ] Enable HTTPS
- [ ] Set up proper logging
- [ ] Configure rate limits appropriately
- [ ] Secure API keys

### Environment Variables

```env
# Production
GOOGLE_API_KEY=<production-key>
DATABASE_URL=postgresql://user:pass@host/db
CORS_ORIGINS=https://your-domain.com
```

---

## Contributing

1. Create a feature branch
2. Make changes with tests
3. Run test suite
4. Submit pull request

### Commit Message Format

```
type(scope): description

- feat: New feature
- fix: Bug fix
- docs: Documentation
- refactor: Code refactoring
- test: Adding tests
```

Example:
```
feat(export): add support for nested tables
```
