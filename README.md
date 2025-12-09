# Document Digital Copy POC

A high-fidelity document processing system for **DOCX** and **Excel (XLSX)** files with AI-powered editing.

```
┌─────────────────────────────────────────────────────────────────────┐
│  📄 DOCX / 📊 Excel  ──►  Parse  ──►  Edit  ──►  Export            │
│                                                                     │
│  ✓ Full Structural Fidelity   ✓ Zero Format Loss   ✓ AI Editing   │
└─────────────────────────────────────────────────────────────────────┘
```

See:
- `docs/HOW_IT_WORKS.md` for a full stage-by-stage explanation
- `docs/ARCHITECTURE.md` for detailed architecture diagrams

## Features

### 📄 DOCX Engine

| Feature | Description |
|---------|-------------|
| **Parsing** | Paragraphs, tables, drawings, checkboxes, dropdowns |
| **Live Preview** | Real-time document preview with formatting |
| **Bidirectional Selection** | Click blocks or preview elements to sync |
| **AI Editing** | LangGraph + Gemini powered text transformations |
| **Form Controls** | Edit checkboxes and dropdowns |
| **High-Fidelity Export** | Preserve formatting, merges, borders, colors |

### 📊 Excel Engine

| Feature | Description |
|---------|-------------|
| **Parsing** | Cells, formulas, styles, merged cells, multiple sheets |
| **Complex Elements** | Conditional formatting (117+ rules), freeze panes, dropdowns |
| **Form Controls** | Checkboxes, radio buttons, spinners from VML |
| **Defined Names** | Named ranges, print areas, built-in names |
| **Tables** | Structured tables (ListObjects) with styles |
| **High-Fidelity Export** | 100% structural preservation via byte-copy strategy |

---

## Quick Start

### Prerequisites

- Python 3.10+
- Node.js 18+
- Google Gemini API key (optional, for AI features)

### Installation

```bash
# Clone and setup
git clone <repository>
cd poc-2

# Backend
python -m venv .venv
.venv\Scripts\Activate.ps1  # Windows
pip install -r requirements.txt

# Frontend
cd web && npm install && cd ..

# Environment
echo "GOOGLE_API_KEY=your-key" > .env
```

### Run

```bash
# Terminal 1: Backend
uvicorn main:app --reload --port 8000

# Terminal 2: Frontend
cd web && npm run dev
```

Open http://localhost:3000

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     FRONTEND (Next.js)                          │
│  ┌───────────────────────────────────────────────────────┐      │
│  │  [📄 DOCX] [📊 Excel]  ◄── Toggle Pills                │      │
│  └───────────────────────────────────────────────────────┘      │
│  ┌─────────┐    ┌─────────┐    ┌─────────┐                      │
│  │ Blocks  │◄──►│ Editor  │◄──►│ Preview │                      │
│  └─────────┘    └─────────┘    └─────────┘                      │
└──────────────────────────┬──────────────────────────────────────┘
                           │ REST API
┌──────────────────────────┴──────────────────────────────────────┐
│                     BACKEND (FastAPI)                            │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  │
│  │ Document Engine │  │  Excel Engine   │  │    AI Agent     │  │
│  │ (DOCX ↔ JSON)   │  │ (XLSX ↔ JSON)   │  │ (LangGraph)     │  │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
```

### Key Components

| Component | Location | Purpose |
|-----------|----------|---------|
| FastAPI App | `main.py` | Application entry |
| Document API | `api/routes/documents.py` | DOCX CRUD, export, AI edit |
| **Spreadsheet API** | `api/routes/spreadsheets.py` | Excel CRUD, export |
| Document Engine | `services/document_engine.py` | DOCX↔JSON conversion |
| **Excel Engine** | `services/excel_engine/` | XLSX↔JSON conversion |
| Edit Service | `services/document_edit_service.py` | AI edit orchestration |
| AI Agent | `services/ai_agent.py` | LangGraph + Gemini |
| AI Config | `services/ai_config.py` | Centralized AI settings |
| Schemas | `models/schemas.py` | DOCX Pydantic models |
| **Excel Schemas** | `services/excel_engine/schemas.py` | Excel Pydantic models |
| Frontend | `web/src/app/page.tsx` | React UI with mode toggle |

---

## API Endpoints

### 📄 DOCX Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/documents/` | Upload DOCX |
| `GET` | `/documents/{id}` | Get document |
| `PUT` | `/documents/{id}` | Update document |
| `POST` | `/documents/{id}/export/file` | Export DOCX |
| `POST` | `/documents/{id}/ai-edit` | AI-powered edit |
| `POST` | `/documents/{id}/checkbox` | Update checkbox |
| `POST` | `/documents/{id}/dropdown` | Update dropdown |

### 📊 Excel Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/spreadsheets/` | Upload XLSX |
| `GET` | `/spreadsheets/{id}` | Get spreadsheet |
| `PUT` | `/spreadsheets/{id}` | Update spreadsheet |
| `POST` | `/spreadsheets/{id}/cell` | Edit single cell |
| `POST` | `/spreadsheets/{id}/cells` | Edit multiple cells |
| `POST` | `/spreadsheets/{id}/export/file` | Export XLSX |

See [docs/API.md](docs/API.md) for full reference.

---

## Document Model

```json
{
  "id": "document.docx",
  "blocks": [
    {"type": "paragraph", "id": "p-0", "runs": [...]},
    {"type": "table", "id": "tbl-0", "rows": [...]},
    {"type": "drawing", "id": "drawing-0", ...}
  ],
  "checkboxes": [{"id": "checkbox-1", "checked": false}],
  "dropdowns": [{"id": "dropdown-1", "selected": "Option A"}]
}
```

See [docs/SCHEMAS.md](docs/SCHEMAS.md) for full schema documentation.

---

## Testing

```bash
# Run all tests
pytest tests/ -v

# Run DOCX fidelity tests
pytest tests/test_fidelity.py -v

# Run Excel engine tests
python tests/test_excel_engine.py
```

### 📄 DOCX Test Results (sample test document)

| Metric | Value |
|--------|-------|
| Blocks | 832 (711 paragraphs, 117 tables, 4 drawings) |
| Checkboxes | 2,973 |
| Dropdowns | 1,050 |
| Text Segments | 12,147 |
| Roundtrip Fidelity | ✓ Pass |

### 📊 Excel Test Results (excel_test.XLSX)

| Metric | Value (example, based on bundled sample) |
|--------|------------------------------------------|
| Sheets | 13 |
| Total Cells | 576,000+ |
| Merged Cells | 8,700+ |
| Conditional Formatting | 117 rules |
| Freeze Panes | 4 sheets |
| Defined Names | 8 |
| Roundtrip Fidelity | ✓ Preserved for sample workbook |

---

## Project Structure

```
poc-2/
├── main.py                 # FastAPI entry point
├── requirements.txt        # Python dependencies
├── api/routes/
│   ├── documents.py        # DOCX API endpoints
│   └── spreadsheets.py     # Excel API endpoints
├── models/schemas.py       # DOCX data models
├── services/
│   ├── document_engine.py  # DOCX processing
│   ├── excel_engine/       # Excel processing
│   │   ├── schemas.py      # 20+ Pydantic models
│   │   ├── parser.py       # XLSX → JSON (~1500 lines)
│   │   └── writer.py       # JSON → XLSX
│   ├── document_edit_service.py
│   ├── ai_agent.py         # LangGraph + Gemini
│   ├── ai_config.py
│   └── db.py
├── middleware/             # Rate limiting
├── tests/
│   ├── docx/               # DOCX-specific tests
│   └── excel/              # Excel-specific tests
├── web/                    # Next.js frontend
├── data/                   # Uploads, exports, DB
└── docs/                   # Documentation
```

---

## Documentation

| Document | Description |
|----------|-------------|
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | System design, data flows, diagrams |
| [API.md](docs/API.md) | Complete API reference |
| [SCHEMAS.md](docs/SCHEMAS.md) | Data model documentation |
| [DEVELOPMENT.md](docs/DEVELOPMENT.md) | Development guide |
| [quickstart.md](quickstart.md) | Step-by-step setup |

---

## Technology Stack

| Layer | Technology |
|-------|------------|
| Frontend | Next.js 14, React 18, TypeScript, TailwindCSS |
| Backend | FastAPI, Python 3.10+, Pydantic 2 |
| AI | LangGraph, Google Gemini |
| Database | SQLite, SQLAlchemy |

---

## Known Limitations

### DOCX
- **Export focuses on text** - Layout/structure changes not supported
- **Drawings are placeholders** - Vector graphics not rendered

### Excel
- **Large file UI performance** - Sheets with 100k+ cells limited in preview
- **AI editing not yet wired** - Manual cell editing only

### General
- **Rate limiting is in-memory** - Not distributed
- **Eval dashboard is temporary** - No persistence

---

## License

MIT License
