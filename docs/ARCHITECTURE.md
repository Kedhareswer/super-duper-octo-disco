# Architecture Documentation

> Complete technical architecture of the DiligenceVault Document Processing System

This document explains the **system design**, **component relationships**, and **data flows** in detail. It covers why each architectural decision was made and how components interact.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Layer Architecture](#layer-architecture)
3. [Component Details](#component-details)
4. [Data Flow Diagrams](#data-flow-diagrams)
5. [AI Agent Architecture](#ai-agent-architecture)
6. [Processing Pipelines](#processing-pipelines)
7. [Technology Decisions](#technology-decisions)
8. [Security Considerations](#security-considerations)

---

## System Overview

The Document Digital Copy POC is a full-stack application that enables high-fidelity **DOCX** and **Excel** document processing with AI-powered editing capabilities.

### Core Design Principles

| Principle | Implementation | Why |
|-----------|----------------|-----|
| **Document Fidelity** | Byte-copy preservation, XML reference tracking | Users expect exported documents to look identical to originals |
| **Separation of Concerns** | Distinct engines for DOCX and Excel | Each format has unique requirements and optimizations |
| **JSON Intermediate** | All documents converted to structured JSON | Enables editing, AI processing, and frontend display |
| **Stateless API** | REST endpoints, no server sessions | Scalability and simplicity |
| **Fail-Safe AI** | Guardrails, validation, retry logic | AI output must be safe and relevant |

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                              USER INTERFACE                                  │
│                         (Next.js / React / TypeScript)                       │
│  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐                      │
│  │   Blocks    │    │   Editor    │    │   Preview   │                      │
│  │   Panel     │◄──►│   Panel     │◄──►│   Panel     │                      │
│  └─────────────┘    └─────────────┘    └─────────────┘                      │
│         │                  │                  │                              │
│         └──────────────────┼──────────────────┘                              │
│                            ▼                                                 │
│                   Bidirectional Sync                                         │
└────────────────────────────┬────────────────────────────────────────────────┘
                             │ HTTP/REST
                             ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                           FASTAPI BACKEND                                    │
│  ┌──────────────────────────────────────────────────────────────────────┐   │
│  │                         API LAYER                                     │   │
│  │  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐       │   │
│  │  │  /documents/*   │  │ /spreadsheets/* │  │    /evals/*     │       │   │
│  │  │  - Upload DOCX  │  │  - Upload XLSX  │  │  - Evaluate     │       │   │
│  │  │  - Get/Update   │  │  - Get/Update   │  │  - Test Suite   │       │   │
│  │  │  - Export       │  │  - Edit Cells   │  │  - Dashboard    │       │   │
│  │  │  - AI Edit      │  │  - Export       │  │                 │       │   │
│  │  │  - Checkbox     │  │                 │  │                 │       │   │
│  │  │  - Dropdown     │  │                 │  │                 │       │   │
│  │  └─────────────────┘  └─────────────────┘  └─────────────────┘       │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                            │                                                 │
│  ┌─────────────────────────┼────────────────────────────────────────────┐   │
│  │                    SERVICE LAYER                                      │   │
│  │  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐       │   │
│  │  │ Document Engine │  │  Excel Engine   │  │   Database      │       │   │
│  │  │ - DOCX→JSON     │  │ - XLSX→JSON     │  │ - SQLite        │       │   │
│  │  │ - JSON→DOCX     │  │ - JSON→XLSX     │  │ - SQLAlchemy    │       │   │
│  │  │ - Validation    │  │ - 100% Fidelity │  │                 │       │   │
│  │  └─────────────────┘  └─────────────────┘  └─────────────────┘       │   │
│  │                              │                                        │   │
│  │  ┌─────────────────┐  ┌─────────────────┐                            │   │
│  │  │  Edit Service   │  │    AI Agent     │                            │   │
│  │  │ - Orchestration │  │ - LangGraph     │                            │   │
│  │  │ - Baked-in Evals│  │ - Gemini        │                            │   │
│  │  └─────────────────┘  └─────────────────┘                            │   │
│  │                              │                                        │   │
│  │                    ┌─────────┴─────────┐                             │   │
│  │                    │     AI Config     │                             │   │
│  │                    │ - Centralized env │                             │   │
│  │                    │ - Provider select │                             │   │
│  │                    └───────────────────┘                             │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
│                            │                                                 │
│  ┌─────────────────────────┼────────────────────────────────────────────┐   │
│  │                    MIDDLEWARE                                         │   │
│  │  ┌─────────────────┐  ┌─────────────────┐                            │   │
│  │  │  Rate Limiter   │  │      CORS       │                            │   │
│  │  └─────────────────┘  └─────────────────┘                            │   │
│  └──────────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│                          DATA LAYER                                          │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐              │
│  │   data/uploads  │  │  data/exports   │  │    data/app.db  │              │
│  │   (DOCX files)  │  │  (DOCX files)   │  │   (SQLite DB)   │              │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘              │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Data Flow Diagrams

### 1. Document Upload Flow

```
┌──────────┐     ┌──────────┐     ┌──────────────┐     ┌──────────┐
│  User    │     │ Frontend │     │   Backend    │     │ Database │
└────┬─────┘     └────┬─────┘     └──────┬───────┘     └────┬─────┘
     │                │                   │                  │
     │ Select DOCX    │                   │                  │
     │───────────────►│                   │                  │
     │                │                   │                  │
     │                │ POST /documents/  │                  │
     │                │──────────────────►│                  │
     │                │                   │                  │
     │                │                   │ Save to disk     │
     │                │                   │─────────────────►│
     │                │                   │                  │
     │                │                   │ Parse DOCX       │
     │                │                   │ (docx_to_json)   │
     │                │                   │                  │
     │                │                   │ Store JSON       │
     │                │                   │─────────────────►│
     │                │                   │                  │
     │                │  DocumentJSON     │                  │
     │                │◄──────────────────│                  │
     │                │                   │                  │
     │ Display        │                   │                  │
     │◄───────────────│                   │                  │
     │                │                   │                  │
```

### 2. AI Edit Flow

```
┌──────────┐     ┌──────────┐     ┌──────────────┐     ┌──────────┐
│  User    │     │ Frontend │     │   Backend    │     │  Gemini  │
└────┬─────┘     └────┬─────┘     └──────┬───────┘     └────┬─────┘
     │                │                   │                  │
     │ Select block   │                   │                  │
     │ Enter instruction                  │                  │
     │───────────────►│                   │                  │
     │                │                   │                  │
     │                │ POST /ai-edit     │                  │
     │                │──────────────────►│                  │
     │                │                   │                  │
     │                │                   │ Route delegates  │
     │                │                   │ to EditService   │
     │                │                   │                  │
     │                │                   │ Locate target    │
     │                │                   │ block/cell       │
     │                │                   │                  │
     │                │                   │ Call AI Agent    │
     │                │                   │─────────────────►│
     │                │                   │◄─────────────────│
     │                │                   │                  │
     │                │                   │ Run Evals        │
     │                │                   │ (baked-in)       │
     │                │                   │                  │
     │                │                   │ Apply edit       │
     │                │                   │ Validate doc     │
     │                │                   │                  │
     │                │                   │ Update DB        │
     │                │                   │                  │
     │                │  Updated JSON     │                  │
     │                │◄──────────────────│                  │
     │                │                   │                  │
     │ Display edit   │                   │                  │
     │◄───────────────│                   │                  │
```

### 3. Export Flow

```
┌──────────┐     ┌──────────┐     ┌──────────────┐     ┌──────────┐
│  User    │     │ Frontend │     │   Backend    │     │   Disk   │
└────┬─────┘     └────┬─────┘     └──────┬───────┘     └────┬─────┘
     │                │                   │                  │
     │ Click Export   │                   │                  │
     │───────────────►│                   │                  │
     │                │                   │                  │
     │                │ POST /export/file │                  │
     │                │──────────────────►│                  │
     │                │                   │                  │
     │                │                   │ Load JSON from DB│
     │                │                   │                  │
     │                │                   │ Load base DOCX   │
     │                │                   │◄─────────────────│
     │                │                   │                  │
     │                │                   │ apply_json_to_docx
     │                │                   │ - Patch text     │
     │                │                   │ - Patch checkboxes
     │                │                   │ - Patch dropdowns│
     │                │                   │ - Fix XML decl   │
     │                │                   │                  │
     │                │                   │ Save new DOCX    │
     │                │                   │─────────────────►│
     │                │                   │                  │
     │                │  DOCX file        │                  │
     │                │◄──────────────────│                  │
     │                │                   │                  │
     │ Download       │                   │                  │
     │◄───────────────│                   │                  │
```

---

## Component Architecture

### Backend Components

| Component | Location | Responsibility |
|-----------|----------|----------------|
| **FastAPI App** | `main.py` | Application entry point, middleware setup |
| **Documents Router** | `api/routes/documents.py` | DOCX CRUD, AI edit, export endpoints |
| **Spreadsheets Router** | `api/routes/spreadsheets.py` | Excel CRUD, cell edit, export endpoints |
| **Evals Router** | `api/routes/evals.py` | AI evaluation dashboard APIs |
| **Document Engine** | `services/document_engine.py` | DOCX↔JSON conversion |
| **Excel Engine** | `services/excel_engine/` | XLSX↔JSON conversion (100% fidelity) |
| **Edit Service** | `services/document_edit_service.py` | AI edit orchestration with baked-in evals |
| **AI Agent** | `services/ai_agent.py` | LangGraph + Gemini integration |
| **AI Config** | `services/ai_config.py` | Centralized provider config |
| **Database** | `services/db.py` | SQLAlchemy models and session management |
| **Rate Limiter** | `middleware/rate_limit.py` | Request rate limiting |
| **DOCX Schemas** | `models/schemas.py` | DOCX Pydantic data models |
| **Excel Schemas** | `services/excel_engine/schemas.py` | Excel Pydantic data models (20+ models) |

### Frontend Components

| Component | Location | Responsibility |
|-----------|----------|----------------|
| **Main Page** | `web/src/app/page.tsx` | Document editor UI with DOCX/Excel toggle |
| **Mode Toggle** | `web/src/app/page.tsx` | Toggle pills for switching between DOCX and Excel |
| **DOCX Preview** | `web/src/app/page.tsx` | Renders DOCX paragraphs, tables, drawings |
| **Excel Preview** | `web/src/app/page.tsx` | Renders Excel grid with sheet tabs |
| **Evals Page** | `web/src/app/evals/page.tsx` | AI evaluation dashboard |
| **Layout** | `web/src/app/layout.tsx` | App shell and styling |

---

## State Machine: AI Edit Agent

The AI agent uses LangGraph to implement an editing workflow. The diagrams in this
section represent the *intended* full state machine; the current implementation in
`services/ai_agent.py` uses a simpler linear pipeline:

```text
validate_input → analyze_intent → execute_edit → validate_output → END
```

There is no multi-step clarify/retry loop yet; on validation failure the agent
falls back to the original text and returns validation errors.

```
                    ┌─────────────────┐
                    │     START       │
                    └────────┬────────┘
                             │
                             ▼
                    ┌─────────────────┐
                    │ validate_input  │
                    │   (Guardrails)  │
                    └────────┬────────┘
                             │
                    ┌────────┴────────┐
                    │                 │
              validation         validation
               passed              failed
                    │                 │
                    ▼                 ▼
           ┌─────────────────┐  ┌─────────────────┐
           │ analyze_intent  │  │      END        │
           │   (Gemini)      │  │ (return errors) │
           └────────┬────────┘  └─────────────────┘
                    │
                    ▼
           ┌─────────────────┐
           │  execute_edit   │
           │   (Gemini)      │
           └────────┬────────┘
                    │
                    ▼
           ┌─────────────────┐
           │ validate_output │
           │   (Guardrails)  │
           └────────┬────────┘
                    │
                    ▼
           ┌─────────────────┐
           │      END        │
           │ (return result) │
           └─────────────────┘
```

### State Definition

```python
class EditState(TypedDict):
    # Input
    original_text: str
    instruction: str
    context: str
    
    # Processing
    intent: str          # "formalize", "simplify", "correct", etc.
    validation_passed: bool
    validation_errors: list[str]
    
    # Output
    edited_text: str
    confidence: float    # 0.0 - 1.0
    reasoning: str
```

---

## DOCX Processing Pipeline

### Parsing (DOCX → JSON)

```
┌─────────────────────────────────────────────────────────────────┐
│                         DOCX File                                │
│  ┌─────────────────────────────────────────────────────────┐    │
│  │                    word/document.xml                     │    │
│  │  ┌─────────┐  ┌─────────┐  ┌─────────┐  ┌─────────┐     │    │
│  │  │  w:p    │  │  w:tbl  │  │  w:sdt  │  │w:drawing│     │    │
│  │  │(para)   │  │(table)  │  │(control)│  │ (image) │     │    │
│  │  └────┬────┘  └────┬────┘  └────┬────┘  └────┬────┘     │    │
│  └───────┼────────────┼────────────┼────────────┼──────────┘    │
└──────────┼────────────┼────────────┼────────────┼───────────────┘
           │            │            │            │
           ▼            ▼            ▼            ▼
    ┌──────────┐ ┌──────────┐ ┌──────────┐
    │Paragraph │ │  Table   │ │ Drawing  │
    │  Block   │ │  Block   │ │  Block   │
    │ (+ inline│ │ (+ inline│ └──────────┘
    │ controls)│ │ controls)│
    └──────────┘ └──────────┘
           │            │
           └────────────┴────────────────────────┘
                                │
                                ▼
                       ┌─────────────────────────┐
                       │     DocumentJSON        │
                       │  - blocks[] (with       │
                       │    inline CheckboxRun   │
                       │    and DropdownRun)     │
                       │  - checkboxes[] (legacy)│
                       │  - dropdowns[] (legacy) │
                       └─────────────────────────┘
```

### Export (JSON → DOCX)

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  DocumentJSON   │     │   Base DOCX     │     │  Output DOCX    │
│  (with edits)   │     │   (original)    │     │  (modified)     │
└────────┬────────┘     └────────┬────────┘     └────────┬────────┘
         │                       │                       │
         │                       │                       │
         ▼                       ▼                       │
    ┌─────────────────────────────────────┐              │
    │         apply_json_to_docx          │              │
    │  1. Load base document.xml          │              │
    │  2. Patch paragraph runs            │──────────────┘
    │  3. Patch table cell text           │
    │  4. Patch checkbox states           │
    │  5. Patch dropdown selections       │
    │  6. Fix XML declaration             │
    │  7. Rebuild DOCX archive            │
    └─────────────────────────────────────┘
```

---

## Technology Stack

| Layer | Technology | Version |
|-------|------------|---------|
| **Frontend** | Next.js | 16.0.5 |
| | React | 19.2.0 |
| | TypeScript | 5.x |
| | TailwindCSS | 4.x |
| **Backend** | FastAPI | 0.104+ |
| | Python | 3.10+ |
| | Pydantic | 2.0+ |
| | SQLAlchemy | 2.0+ |
| **AI** | LangGraph | 0.0.40+ |
| | Google Gemini | 0.3.0+ |
| **Database** | SQLite | (bundled) |

---

## Security Considerations

### Rate Limiting

```
┌─────────────────────────────────────────────────────────────┐
│                    Rate Limit Configuration                  │
├─────────────────────────────────┬───────────────────────────┤
│ Limit Type                      │ Value                     │
├─────────────────────────────────┼───────────────────────────┤
│ Requests per minute (general)   │ 60                        │
│ Requests per hour (general)     │ 1000                      │
│ Requests per minute (AI)        │ 20                        │
│ Requests per hour (AI)          │ 200                       │
│ Burst limit (per second)        │ 15                        │
└─────────────────────────────────┴───────────────────────────┘
```

### Guardrails

| Check | Location | Purpose |
|-------|----------|---------|
| Input length | `Guardrails.validate_input` | Prevent DoS via large inputs |
| Blocked patterns | `Guardrails.validate_input` | Block sensitive data (SSN, CC) |
| Output length | `Guardrails.validate_output` | Prevent runaway generation |
| Hallucination detection | `Guardrails.validate_output` | Detect AI self-references |
| Content preservation | `Guardrails.validate_output` | Ensure meaningful output |

---

## File Structure

```
poc-2/
├── main.py                 # FastAPI application entry point
├── requirements.txt        # Python dependencies
├── .env                    # Environment variables (API keys)
│
├── api/                    # API layer
│   └── routes/
│       ├── documents.py    # DOCX CRUD, AI edit, export
│       ├── spreadsheets.py # Excel CRUD, cell edit, export
│       └── evals.py        # AI evaluation endpoints
│
├── models/                 # Data models
│   └── schemas.py          # DOCX Pydantic schemas
│
├── services/               # Business logic
│   ├── document_engine.py  # DOCX↔JSON conversion
│   ├── excel_engine/       # Excel processing (isolated module)
│   │   ├── __init__.py     # Public API exports
│   │   ├── schemas.py      # 20+ Pydantic models for Excel
│   │   ├── parser.py       # XLSX→JSON (~1500 lines)
│   │   └── writer.py       # JSON→XLSX (byte-copy strategy)
│   ├── document_edit_service.py  # AI edit orchestration
│   ├── ai_agent.py         # LangGraph + Gemini AI agent
│   ├── ai_config.py        # Centralized AI configuration
│   └── db.py               # SQLAlchemy database layer
│
├── middleware/             # HTTP middleware
│   └── rate_limit.py       # Rate limiting
│
├── data/                   # Runtime data
│   ├── uploads/            # Uploaded DOCX/XLSX files
│   ├── exports/            # Exported DOCX/XLSX files
│   └── app.db              # SQLite database
│
├── tests/                  # Test suite
│   ├── test_backend_e2e.py # End-to-end backend tests
│   ├── test_fidelity.py    # DOCX fidelity tests
│   ├── test_excel_engine.py# Excel fidelity tests
│   └── test_export_roundtrip.py
│
├── debug/                  # Debug utilities (not for production)
│
├── web/                    # Next.js frontend
│   ├── src/app/
│   │   ├── page.tsx        # Main editor (DOCX/Excel toggle)
│   │   ├── evals/page.tsx  # AI evaluation dashboard
│   │   └── layout.tsx      # App layout
│   └── package.json
│
└── docs/                   # Documentation
    ├── ARCHITECTURE.md     # This file
    ├── API.md              # API reference
    ├── SCHEMAS.md          # Data model documentation
    └── DEVELOPMENT.md      # Development guide
```
