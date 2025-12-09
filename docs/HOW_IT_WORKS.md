# How It Works - Complete Technical Guide

> This document explains the **complete processing pipeline** from file upload to export. It covers every stage, the technology used, and why each design decision was made.

---

## Table of Contents

1. [Overview: The Complete Pipeline](#1-overview-the-complete-pipeline)
2. [Stage 1: File Upload](#2-stage-1-file-upload)
3. [Stage 2: Document Parsing](#3-stage-2-document-parsing)
4. [Stage 3: JSON Representation](#4-stage-3-json-representation)
5. [Stage 4: Frontend Display](#5-stage-4-frontend-display)
6. [Stage 5: Editing (Manual & AI)](#6-stage-5-editing-manual--ai)
7. [Stage 6: Document Export](#7-stage-6-document-export)
8. [DOCX Processing Deep Dive](#8-docx-processing-deep-dive)
9. [Excel Processing Deep Dive](#9-excel-processing-deep-dive)
10. [AI Agent Pipeline](#10-ai-agent-pipeline)

---

## 1. Overview: The Complete Pipeline

The system processes documents through a **6-stage pipeline**:

```mermaid
flowchart TB
    subgraph Stage1["Stage 1: Upload"]
        Upload[User uploads file]
        Save[Save to disk]
    end
    
    subgraph Stage2["Stage 2: Parse"]
        Unzip[Unzip OOXML package]
        ParseXML[Parse XML files]
        Extract[Extract content & structure]
    end
    
    subgraph Stage3["Stage 3: JSON Model"]
        BuildJSON[Build JSON representation]
        StoreDB[Store in database]
    end
    
    subgraph Stage4["Stage 4: Display"]
        SendFrontend[Send to frontend]
        RenderUI[Render in editor]
    end
    
    subgraph Stage5["Stage 5: Edit"]
        UserEdit[User makes changes]
        AIEdit[AI processes text]
        UpdateJSON[Update JSON model]
    end
    
    subgraph Stage6["Stage 6: Export"]
        ApplyChanges[Apply changes to XML]
        Repackage[Repackage OOXML]
        Download[User downloads file]
    end
    
    Stage1 --> Stage2 --> Stage3 --> Stage4 --> Stage5 --> Stage6
```

### Why This Architecture?

| Design Decision | Reason |
|-----------------|--------|
| **JSON intermediate format** | Enables frontend editing, AI processing, and database storage without touching raw XML |
| **Preserve original file** | Allows byte-copy export for unchanged parts, ensuring 100% fidelity |
| **Separate engines** | DOCX and XLSX have different structures; optimized handling for each |
| **XML reference tracking** | Enables precise mapping between JSON edits and XML locations |

---

## 2. Stage 1: File Upload

### What Happens

1. User selects a file in the browser
2. Frontend sends file via `multipart/form-data` POST request
3. Backend receives and validates the file
4. File is saved to disk with a unique ID prefix

```mermaid
sequenceDiagram
    participant User
    participant Browser
    participant FastAPI
    participant FileSystem
    
    User->>Browser: Select file
    Browser->>FastAPI: POST /documents/ or /spreadsheets/
    FastAPI->>FastAPI: Validate file type
    FastAPI->>FileSystem: Save to data/uploads/{type}/
    FastAPI->>FastAPI: Generate unique ID
    FastAPI-->>Browser: Return document ID
```

### Technology Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| Frontend upload | `<input type="file">` | File selection |
| HTTP transport | `fetch()` with FormData | Multipart upload |
| Backend handler | FastAPI `UploadFile` | Async file handling |
| File storage | Python `pathlib` | Cross-platform paths |

### Code Location

```
api/routes/documents.py    → upload_document()
api/routes/spreadsheets.py → upload_spreadsheet()
```

### File Naming Convention

```
{8-char-uuid}_{original-filename}
Example: 24adb5ab_excel_test.XLSX
```

**Why?** Prevents filename collisions while preserving the original name for user reference.

---

## 3. Stage 2: Document Parsing

### What Happens

The parser **unpacks the OOXML package** and **extracts structured content**.

#### Understanding OOXML

Both DOCX and XLSX files are **ZIP archives** containing XML files:

```
document.docx (unzipped)
├── [Content_Types].xml      # File type declarations
├── _rels/
│   └── .rels               # Package relationships
├── word/
│   ├── document.xml        # Main content ← WE PARSE THIS
│   ├── styles.xml          # Style definitions
│   └── _rels/
│       └── document.xml.rels
└── docProps/
    └── core.xml            # Metadata
```

```
spreadsheet.xlsx (unzipped)
├── [Content_Types].xml
├── xl/
│   ├── workbook.xml        # Workbook structure
│   ├── sharedStrings.xml   # Shared string table ← IMPORTANT
│   ├── styles.xml          # Cell styles
│   └── worksheets/
│       ├── sheet1.xml      # Sheet content ← WE PARSE THESE
│       └── sheet2.xml
└── _rels/
```

### DOCX Parsing Flow

```mermaid
flowchart LR
    subgraph Input
        DOCX[DOCX File]
    end
    
    subgraph Unpack
        ZIP[zipfile.ZipFile]
        XML[word/document.xml]
    end
    
    subgraph Parse
        ET[ElementTree.parse]
        Body[Find w:body]
    end
    
    subgraph Extract
        Paragraphs[Extract paragraphs]
        Tables[Extract tables]
        Drawings[Extract drawings]
        Controls[Extract checkboxes/dropdowns]
    end
    
    subgraph Output
        JSON[DocumentJSON]
    end
    
    DOCX --> ZIP --> XML --> ET --> Body
    Body --> Paragraphs --> JSON
    Body --> Tables --> JSON
    Body --> Drawings --> JSON
    Body --> Controls --> JSON
```

### Excel Parsing Flow

```mermaid
flowchart LR
    subgraph Input
        XLSX[XLSX File]
    end
    
    subgraph Unpack
        ZIP[zipfile.ZipFile]
        Workbook[xl/workbook.xml]
        Sheets[xl/worksheets/*.xml]
        Strings[xl/sharedStrings.xml]
    end
    
    subgraph Parse
        ParseWB[Parse workbook structure]
        ParseSS[Parse shared strings]
        ParseSheet[Parse each sheet]
    end
    
    subgraph Extract
        Cells[Extract cells]
        Merges[Extract merged ranges]
        Validations[Extract data validations]
        Images[Extract images]
    end
    
    subgraph Output
        JSON[ExcelWorkbookJSON]
    end
    
    XLSX --> ZIP
    ZIP --> Workbook --> ParseWB
    ZIP --> Strings --> ParseSS
    ZIP --> Sheets --> ParseSheet
    ParseSheet --> Cells --> JSON
    ParseSheet --> Merges --> JSON
    ParseSheet --> Validations --> JSON
    ParseSheet --> Images --> JSON
```

### Technology Stack

| Component | Technology | Why This Choice |
|-----------|------------|-----------------|
| ZIP handling | `zipfile` (stdlib) | No external dependencies, reliable |
| XML parsing | `xml.etree.ElementTree` (stdlib) | Fast, memory-efficient, namespace-aware |
| Data models | Pydantic v2 | Validation, serialization, type safety |

### XML Namespace Handling

OOXML uses many XML namespaces. The parser registers them for proper element lookup:

```python
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    # ... more namespaces
}
```

**Why?** Without namespace registration, `element.find("w:p")` would fail because ElementTree doesn't know what `w:` means.

---

## 4. Stage 3: JSON Representation

### What Happens

The parsed content is converted to a **structured JSON model** that:
- Preserves document hierarchy
- Maintains formatting information
- Includes XML references for export
- Enables programmatic editing

### DOCX JSON Structure

```mermaid
classDiagram
    class DocumentJSON {
        +str id
        +str title
        +List~Block~ blocks
        +List~CheckboxField~ checkboxes
        +List~DropdownField~ dropdowns
    }
    
    class ParagraphBlock {
        +str type = "paragraph"
        +str id
        +str xml_ref
        +str style_name
        +List~Run~ runs
    }
    
    class TableBlock {
        +str type = "table"
        +str id
        +str xml_ref
        +List~TableRow~ rows
    }
    
    class Run {
        +str id
        +str xml_ref
        +str text
        +bool bold
        +bool italic
    }
    
    DocumentJSON --> ParagraphBlock
    DocumentJSON --> TableBlock
    ParagraphBlock --> Run
```

### Excel JSON Structure

```mermaid
classDiagram
    class ExcelWorkbookJSON {
        +str id
        +List~ExcelSheetJSON~ sheets
        +List~SharedStringItem~ shared_strings
    }
    
    class ExcelSheetJSON {
        +str name
        +List~ExcelCellJSON~ cells
        +List~MergedCellRange~ merged_cells
        +List~DataValidationRule~ data_validations
    }
    
    class ExcelCellJSON {
        +str ref
        +Any value
        +CellDataType data_type
        +bool dirty
        +CellStyle style
    }
    
    ExcelWorkbookJSON --> ExcelSheetJSON
    ExcelSheetJSON --> ExcelCellJSON
```

### XML Reference System

Every editable element has an `xml_ref` that maps back to its location in the XML:

| Element Type | XML Reference Format | Example |
|--------------|---------------------|---------|
| Paragraph | `p[index]` | `p[0]`, `p[5]` |
| Table | `tbl[index]` | `tbl[0]` |
| Table Cell | `tbl[t]/tr[r]/tc[c]` | `tbl[0]/tr[1]/tc[2]` |
| Run | `p[index]/r[index]` | `p[0]/r[0]` |
| Excel Cell | `{sheet}!{ref}` | `Sheet1!A1` |

**Why XML References?**
- Enable precise location of elements during export
- Allow edits to be applied to exact XML locations
- Support undo/redo by tracking change locations

---

## 5. Stage 4: Frontend Display

### What Happens

The JSON model is sent to the frontend and rendered in a **three-panel editor**:

```mermaid
flowchart LR
    subgraph Backend
        JSON[JSON Model]
    end
    
    subgraph Frontend
        subgraph Panels
            Blocks[Blocks Panel<br/>Navigation tree]
            Editor[Editor Panel<br/>Content editing]
            Preview[Preview Panel<br/>Live preview]
        end
    end
    
    JSON -->|HTTP GET| Blocks
    JSON -->|HTTP GET| Editor
    JSON -->|HTTP GET| Preview
    
    Editor -->|onChange| Blocks
    Editor -->|onChange| Preview
```

### Panel Responsibilities

| Panel | Purpose | Updates When |
|-------|---------|--------------|
| **Blocks** | Shows document structure as a tree | Document loaded, structure changes |
| **Editor** | Allows content editing | User selects a block |
| **Preview** | Shows formatted output | Any content changes |

### Technology Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| Framework | Next.js 14 | Server-side rendering, routing |
| UI Library | React 18 | Component-based UI |
| Styling | TailwindCSS | Utility-first CSS |
| State | React useState/useEffect | Local component state |
| HTTP | fetch API | Backend communication |

---

## 6. Stage 5: Editing (Manual & AI)

### Manual Editing Flow

```mermaid
sequenceDiagram
    participant User
    participant Editor
    participant Backend
    participant Database
    
    User->>Editor: Type new text
    Editor->>Backend: PUT /documents/{id}
    Backend->>Backend: Validate JSON
    Backend->>Database: Update stored JSON
    Backend-->>Editor: Return updated JSON
    Editor->>Editor: Re-render preview
```

### AI Editing Flow (DOCX Only)

```mermaid
sequenceDiagram
    participant User
    participant Frontend
    participant EditService
    participant AIAgent
    participant Gemini
    
    User->>Frontend: Select block + enter instruction
    Frontend->>EditService: POST /documents/{id}/ai-edit
    EditService->>EditService: Locate target block
    EditService->>AIAgent: Process edit request
    AIAgent->>Gemini: Send prompt with context
    Gemini-->>AIAgent: Return transformed text
    AIAgent->>AIAgent: Run guardrails & validation
    AIAgent-->>EditService: Return validated result
    EditService->>EditService: Apply edit to JSON
    EditService-->>Frontend: Return updated document
```

### AI Agent Architecture

The AI agent uses **LangGraph** for workflow orchestration. The diagrams in this
section show the *intended* state machine design; the current implementation in
`services/ai_agent.py` uses a simpler linear workflow:

```text
validate_input → analyze_intent → execute_edit → validate_output → END
```

There is no multi-step clarify/retry loop yet; on output validation failure the
current implementation falls back to the original text and records validation
errors.

```mermaid
stateDiagram-v2
    [*] --> ValidateInput
    ValidateInput --> AnalyzeIntent: Valid
    ValidateInput --> Error: Invalid
    
    AnalyzeIntent --> ExecuteEdit: Clear intent
    AnalyzeIntent --> Clarify: Ambiguous
    
    ExecuteEdit --> ValidateOutput
    ValidateOutput --> ApplyEdit: Pass
    ValidateOutput --> Retry: Fail (< 3 attempts)
    ValidateOutput --> Error: Fail (>= 3 attempts)
    
    Retry --> ExecuteEdit
    ApplyEdit --> [*]
    Error --> [*]
    Clarify --> [*]
```

### Why LangGraph?

| Feature | Benefit |
|---------|---------|
| **State machine** | Clear workflow stages, easy debugging |
| **Conditional edges** | Handle success/failure paths |
| **Retry logic** | Automatic retry on validation failure |
| **Observability** | Built-in tracing and logging |

---

## 7. Stage 6: Document Export

### What Happens

The export process **applies JSON changes back to the original XML** and **repackages the OOXML archive**.

### Export Strategy Comparison

| Aspect | DOCX Engine | Excel Engine |
|--------|-------------|--------------|
| **Strategy** | Full XML modification | Byte-copy with selective updates |
| **Changed files** | `word/document.xml` only | Only modified sheets + sharedStrings |
| **Unchanged files** | Copied from original | Byte-for-byte copy for untouched parts |
| **Namespace handling** | Preserve all declarations | Preserve original root tags |

### DOCX Export Flow

```mermaid
flowchart TB
    subgraph Input
        JSON[Edited JSON]
        Original[Original DOCX]
    end
    
    subgraph Process
        LoadXML[Load document.xml]
        FindElements[Find elements by xml_ref]
        ApplyChanges[Apply text changes]
        Serialize[Serialize back to XML]
    end
    
    subgraph Package
        CopyFiles[Copy unchanged files]
        WriteDoc[Write modified document.xml]
        CreateZip[Create new ZIP]
    end
    
    subgraph Output
        NewDOCX[Exported DOCX]
    end
    
    JSON --> FindElements
    Original --> LoadXML --> FindElements
    FindElements --> ApplyChanges --> Serialize
    Original --> CopyFiles
    Serialize --> WriteDoc
    CopyFiles --> CreateZip
    WriteDoc --> CreateZip
    CreateZip --> NewDOCX
```

### Excel Export Flow

```mermaid
flowchart TB
    subgraph Input
        JSON[Edited JSON with dirty flags]
        Original[Original XLSX]
    end
    
    subgraph Check
        CheckDirty{Any dirty cells?}
    end
    
    subgraph NoDirty[No Changes Path]
        CopyExact[Copy file exactly]
    end
    
    subgraph HasDirty[Has Changes Path]
        IdentifySheets[Identify affected sheets]
        LoadSheetXML[Load sheet XML]
        PreserveRoot[Extract original root tag]
        ModifyCells[Modify cell values]
        SerializeInner[Serialize inner content]
        ReassembleXML[Reassemble with original root]
        UpdateStrings[Update sharedStrings.xml]
    end
    
    subgraph Package
        CopyUnchanged[Copy unchanged files byte-for-byte]
        WriteChanged[Write modified XMLs]
        CreateZip[Create new ZIP]
    end
    
    JSON --> CheckDirty
    CheckDirty -->|No| NoDirty --> CreateZip
    CheckDirty -->|Yes| HasDirty
    Original --> IdentifySheets
    IdentifySheets --> LoadSheetXML --> PreserveRoot
    PreserveRoot --> ModifyCells --> SerializeInner --> ReassembleXML
    ReassembleXML --> WriteChanged
    UpdateStrings --> WriteChanged
    Original --> CopyUnchanged --> CreateZip
    WriteChanged --> CreateZip
```

### Critical: Namespace Preservation

Excel requires **all namespace declarations** to be present, even if they appear "unused":

```xml
<!-- Original root tag - MUST be preserved exactly -->
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           mc:Ignorable="x14ac xr xr2 xr3"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
           xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
           xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
           xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
```

**Why?** The `mc:Ignorable` attribute references these namespaces. If they're missing, Excel reports the file as corrupted.

**Solution in this POC:** Extract the original root tag with regex, serialize only inner
content with ElementTree, then reassemble. This strategy has been validated against the
bundled test workbooks; for arbitrary Excel files, it is a design goal rather than a
formal guarantee.

---

## 8. DOCX Processing Deep Dive

### Document Structure

A DOCX document body contains these element types:

```mermaid
flowchart TB
    Body[w:body]
    
    Body --> P1[w:p - Paragraph]
    Body --> Tbl[w:tbl - Table]
    Body --> P2[w:p - Paragraph]
    Body --> SDT[w:sdt - Content Control]
    
    P1 --> R1[w:r - Run]
    P1 --> R2[w:r - Run]
    R1 --> T1[w:t - Text]
    
    Tbl --> TR1[w:tr - Row]
    Tbl --> TR2[w:tr - Row]
    TR1 --> TC1[w:tc - Cell]
    TR1 --> TC2[w:tc - Cell]
    TC1 --> CP[w:p - Cell Paragraph]
    
    SDT --> SDTPr[w:sdtPr - Properties]
    SDT --> SDTContent[w:sdtContent - Content]
    SDTPr --> Checkbox[w14:checkbox]
    SDTPr --> Dropdown[w:comboBox]
```

### Parsing Algorithm

```python
def parse_document(docx_path: str) -> DocumentJSON:
    # 1. Open ZIP and read document.xml
    with zipfile.ZipFile(docx_path) as zf:
        doc_xml = zf.read("word/document.xml")
    
    # 2. Parse XML
    tree = ET.parse(BytesIO(doc_xml))
    body = tree.find(".//w:body", NS)
    
    # 3. Iterate through body children
    blocks = []
    
    for i, element in enumerate(body):
        if element.tag == f"{{{NS['w']}}}p":
            # Paragraphs may contain inline SDT controls (checkboxes/dropdowns)
            blocks.append(parse_paragraph(element, i))
        elif element.tag == f"{{{NS['w']}}}tbl":
            # Tables may have SDT controls at cell or row level
            blocks.append(parse_table(element, i))
        elif element.tag == f"{{{NS['w']}}}sdt":
            # Body-level SDT - extract content
            blocks.extend(parse_body_sdt(element, i))
    
    # Legacy arrays populated from inline controls for backward compatibility
    checkboxes = doc.get_all_checkboxes()  # Extracts from inline runs
    dropdowns = doc.get_all_dropdowns()    # Extracts from inline runs
    
    return DocumentJSON(blocks=blocks, checkboxes=checkboxes, dropdowns=dropdowns)
```

### Table Cell Merging

DOCX uses `gridSpan` for horizontal merges and `vMerge` for vertical:

| Property | Meaning | JSON Field |
|----------|---------|------------|
| `w:gridSpan val="3"` | Cell spans 3 columns | `col_span: 3` |
| `w:vMerge val="restart"` | Start of vertical merge | `row_span: N` (calculated) |
| `w:vMerge` (no val) | Continuation of merge | Cell omitted |

---

## 9. Excel Processing Deep Dive

### Workbook Structure

```mermaid
flowchart TB
    subgraph Workbook
        WB[xl/workbook.xml]
        SS[xl/sharedStrings.xml]
        Styles[xl/styles.xml]
    end
    
    subgraph Sheets
        S1[xl/worksheets/sheet1.xml]
        S2[xl/worksheets/sheet2.xml]
    end
    
    subgraph Relationships
        Rels[xl/_rels/workbook.xml.rels]
    end
    
    WB --> Rels
    Rels --> S1
    Rels --> S2
    S1 --> SS
    S2 --> SS
```

### Shared Strings Table

Excel stores unique strings in a shared table to reduce file size:

```xml
<!-- xl/sharedStrings.xml -->
<sst count="100" uniqueCount="50">
    <si><t>Hello</t></si>      <!-- index 0 -->
    <si><t>World</t></si>      <!-- index 1 -->
    <si><t>Hello</t></si>      <!-- reuses index 0 -->
</sst>
```

```xml
<!-- xl/worksheets/sheet1.xml -->
<c r="A1" t="s"><v>0</v></c>  <!-- References "Hello" -->
<c r="A2" t="s"><v>1</v></c>  <!-- References "World" -->
```

### Cell Data Types

| Type Code | Meaning | Value Storage |
|-----------|---------|---------------|
| `s` | Shared string | Index into sharedStrings.xml |
| `n` | Number | Inline numeric value |
| `b` | Boolean | 0 or 1 |
| `e` | Error | Error code (#REF!, #VALUE!, etc.) |
| `str` | Inline string | Inline text (rare) |
| (none) | Number | Inline numeric value |

### Data Validation (Dropdowns)

```xml
<dataValidations count="1">
    <dataValidation type="list" sqref="B2:B100">
        <formula1>"Yes,No,Maybe"</formula1>
    </dataValidation>
</dataValidations>
```

---

## 10. AI Agent Pipeline

### Agent State Machine

```mermaid
stateDiagram-v2
    [*] --> Idle
    
    Idle --> Processing: Receive edit request
    
    state Processing {
        [*] --> Validate
        Validate --> Analyze: Valid input
        Validate --> Error: Invalid input
        
        Analyze --> Generate: Intent understood
        Analyze --> Clarify: Ambiguous intent
        
        Generate --> Evaluate: Text generated
        
        Evaluate --> Apply: Passes all checks
        Evaluate --> Retry: Fails checks
        Retry --> Generate: Attempts < 3
        Retry --> Error: Attempts >= 3
    }
    
    Processing --> Complete: Success
    Processing --> Failed: Error
    
    Complete --> Idle
    Failed --> Idle
```

### Guardrails

| Guardrail | Purpose | Action on Failure |
|-----------|---------|-------------------|
| **Length check** | Output not too long/short | Retry with adjusted prompt |
| **Format preservation** | Maintains structure | Retry with format reminder |
| **Content safety** | No harmful content | Reject and return error |
| **Relevance check** | Output matches intent | Retry with clarification |

### Prompt Template

```python
EDIT_PROMPT = """
You are editing a document block. Apply the user's instruction to the text.

CURRENT TEXT:
{current_text}

USER INSTRUCTION:
{instruction}

RULES:
1. Only modify what the instruction asks
2. Preserve formatting markers
3. Keep the same general length unless asked to expand/shorten
4. Return ONLY the edited text, no explanations

EDITED TEXT:
"""
```

---

## Summary

The DiligenceVault Document Processing System achieves **100% document fidelity** through:

1. **Careful XML parsing** that preserves all structural information
2. **JSON intermediate format** that enables editing while maintaining references
3. **Byte-copy export** that only modifies what's necessary
4. **Namespace preservation** that satisfies strict OOXML requirements
5. **AI guardrails** that ensure edits are safe and relevant

Each stage is designed to **minimize data loss** and **maximize compatibility** with Microsoft Office applications.
