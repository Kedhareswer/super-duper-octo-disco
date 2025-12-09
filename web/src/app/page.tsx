"use client";

import { useRef, useState } from "react";

const BACKEND_URL = "http://127.0.0.1:8000";

// =============================================================================
// MODE TOGGLE
// =============================================================================

type DocumentMode = "docx" | "excel";

// =============================================================================
// DOCX TYPES
// =============================================================================

type Run = {
  id: string;
  xml_ref: string;
  text: string;
  bold: boolean;
  italic: boolean;
  color?: string | null;  // Hex color e.g. "FF0000"
};

type ParagraphBlock = {
  type: "paragraph";
  id: string;
  xml_ref: string;
  style_name?: string | null;
  runs: Run[];
};

type CellBorder = {
  style: string;
  width: number;
  color?: string | null;
};

type CellBorders = {
  top?: CellBorder | null;
  bottom?: CellBorder | null;
  left?: CellBorder | null;
  right?: CellBorder | null;
};

// Forward declare TableBlock for recursive type
type CellBlock = ParagraphBlock | TableBlock;

type TableCell = {
  id: string;
  xml_ref: string;
  row_span: number;
  col_span: number;
  background_color?: string | null;
  borders?: CellBorders | null;
  v_merge?: string | null;  // "restart" | "continue" | null
  blocks: CellBlock[];  // Can contain paragraphs and nested tables
};

type TableRow = {
  id: string;
  xml_ref: string;
  cells: TableCell[];
};

type TableBlock = {
  type: "table";
  id: string;
  xml_ref: string;
  rows: TableRow[];
};

type DrawingBlock = {
  type: "drawing";
  id: string;
  xml_ref: string;
  name?: string | null;
  width_inches: number;
  height_inches: number;
  drawing_type: string;
};

type Block = ParagraphBlock | TableBlock | DrawingBlock;

type CheckboxField = {
  id: string;
  xml_ref: string;
  label?: string | null;
  checked: boolean;
};

type DropdownField = {
  id: string;
  xml_ref: string;
  label?: string | null;
  options: string[];
  selected?: string | null;
};

type DocumentJSON = {
  id: string;
  title?: string | null;
  blocks: Block[];
  checkboxes: CheckboxField[];
  dropdowns: DropdownField[];
};

const normalizeDocument = (raw: DocumentJSON): DocumentJSON => {
  return {
    ...raw,
    blocks: raw.blocks ?? [],
    checkboxes: raw.checkboxes ?? [],
    dropdowns: raw.dropdowns ?? [],
  };
};

// =============================================================================
// EXCEL TYPES
// =============================================================================

type ExcelCellBorders = {
  left?: string | null;
  right?: string | null;
  top?: string | null;
  bottom?: string | null;
};

type ExcelCellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string | null;
  bg_color?: string | null;
  font_size?: number | null;
  pattern?: string | null;
  h_align?: string | null;
  v_align?: string | null;
  wrap?: boolean;
  borders?: ExcelCellBorders | null;
};

type ExcelCellDropdown = {
  validation_id: string;
  options: string[];
};

type ExcelCellCheckbox = {
  control_id: string;
  checked: boolean;
};

type ExcelCell = {
  id: string;
  ref: string;
  row: number;
  col: number;
  value: string | number | boolean | null;
  formula?: string | null;
  is_merged: boolean;
  merge_range?: string | null;
  is_merge_origin: boolean;
  style?: ExcelCellStyle | null;
  dropdown?: ExcelCellDropdown | null;
  checkbox?: ExcelCellCheckbox | null;
};

type ExcelMergedCell = {
  ref: string;
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
};

type ExcelValidation = {
  id: string;
  sqref: string;
  type: string;
  options: string[];
};

type ExcelFormControl = {
  id: string;
  type: string;
  checked?: boolean | null;
  linked_cell?: string | null;
};

type ExcelSheet = {
  id: string;
  name: string;
  index: number;
  is_hidden: boolean;
  dimension?: string | null;
  cells: ExcelCell[];
  cell_count: number;
  merged_cells: ExcelMergedCell[];
  data_validations: ExcelValidation[];
  conditional_formatting_count: number;
  form_controls: ExcelFormControl[];
  images_count: number;
  comments_count: number;
  hyperlinks_count: number;
  freeze_pane?: { rows: number; cols: number } | null;
  zoom: number;
};

type ExcelDefinedName = {
  name: string;
  value: string;
  is_builtin: boolean;
};

type SpreadsheetJSON = {
  id: string;
  filename?: string | null;
  sheets: ExcelSheet[];
  active_sheet_index: number;
  defined_names: ExcelDefinedName[];
  metadata?: {
    created?: string | null;
    modified?: string | null;
    creator?: string | null;
  };
};

// =============================================================================
// TOGGLE PILL COMPONENT
// =============================================================================

function ModePill({ 
  mode, 
  currentMode, 
  onClick, 
  label 
}: { 
  mode: DocumentMode; 
  currentMode: DocumentMode; 
  onClick: () => void; 
  label: string;
}) {
  const isActive = mode === currentMode;
  return (
    <button
      onClick={onClick}
      className={`px-4 py-2 text-sm font-medium rounded-full transition-all ${
        isActive
          ? "bg-zinc-900 text-white shadow-md"
          : "bg-zinc-100 text-zinc-600 hover:bg-zinc-200"
      }`}
    >
      {label}
    </button>
  );
}


export default function Home() {
  // Mode toggle
  const [mode, setMode] = useState<DocumentMode>("docx");
  
  // DOCX state
  const [documentId, setDocumentId] = useState<string | null>(null);
  const [doc, setDoc] = useState<DocumentJSON | null>(null);
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null);
  const [selectedCellId, setSelectedCellId] = useState<string | null>(null);
  
  // Excel state
  const [spreadsheetId, setSpreadsheetId] = useState<string | null>(null);
  const [spreadsheet, setSpreadsheet] = useState<SpreadsheetJSON | null>(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState<number>(0);
  const [selectedExcelCell, setSelectedExcelCell] = useState<string | null>(null);
  const [excelCellEditValue, setExcelCellEditValue] = useState<string>("");
  const [isSavingExcelCell, setIsSavingExcelCell] = useState(false);
  
  // Shared state
  const [instruction, setInstruction] = useState<string>("");
  const [isUploading, setIsUploading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportPath, setExportPath] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const previewRefs = useRef<Record<string, HTMLElement | null>>({});
  const blockListRefs = useRef<Record<string, HTMLElement | null>>({});

  const uploadDocxFile = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);

    setIsUploading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/documents/`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        throw new Error(`Upload failed: ${res.status}`);
      }
      const data = (await res.json()) as DocumentJSON;
      setDocumentId(data.id);
      setDoc(normalizeDocument(data));
      setSelectedBlockId(null);
      setExportPath(null);
    } catch (err) {
      console.error(err);
      alert("Upload failed. Check console for details.");
    } finally {
      setIsUploading(false);
    }
  };

  const uploadExcelFile = async (file: File) => {
    const formData = new FormData();
    formData.append("file", file);

    setIsUploading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/spreadsheets/`, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        throw new Error(`Upload failed: ${res.status}`);
      }
      const data = (await res.json()) as SpreadsheetJSON;
      setSpreadsheetId(data.id);
      setSpreadsheet(data);
      setActiveSheetIndex(data.active_sheet_index || 0);
      setSelectedExcelCell(null);
      setExportPath(null);
    } catch (err) {
      console.error(err);
      alert("Upload failed. Check console for details.");
    } finally {
      setIsUploading(false);
    }
  };

  const uploadFile = async (file: File) => {
    if (mode === "excel") {
      await uploadExcelFile(file);
    } else {
      await uploadDocxFile(file);
    }
  };

  const getCellText = (cell: TableCell): string =>
    cell.blocks
      .filter((b): b is ParagraphBlock => b.type === "paragraph")
      .flatMap((p) => p.runs.map((r: Run) => r.text))
      .join(" ")
      .trim();

  const renderPreview = () => {
    if (!doc) {
      return (
        <p className="text-sm text-zinc-500">
          Upload a DOCX to see the preview.
        </p>
      );
    }

    return (
      <div className="flex h-full flex-col gap-2 overflow-auto text-sm">
        {doc.blocks.map((block) => {
          if (block.type === "paragraph") {
            const para = block as ParagraphBlock;
            const isSelected = selectedBlockId === para.id;
            // Render runs with styling
            const paraContent = para.runs.map((r, ri) => {
              const style: React.CSSProperties = {};
              if (r.color) {
                style.color = `#${r.color}`;
              }
              if (r.bold) {
                style.fontWeight = "bold";
              }
              if (r.italic) {
                style.fontStyle = "italic";
              }
              return (
                <span key={`${para.id}-${ri}`} style={style}>
                  {r.text}
                </span>
              );
            });
            return (
              <p
                key={para.id}
                ref={(el) => {
                  if (el) {
                    previewRefs.current[para.id] = el;
                  }
                }}
                onClick={() => {
                  setSelectedBlockId(para.id);
                  setSelectedCellId(null);
                  // Scroll to block in left panel
                  const blockEl = blockListRefs.current[para.id];
                  if (blockEl) {
                    blockEl.scrollIntoView({ behavior: "smooth", block: "center" });
                  }
                }}
                className={`whitespace-pre-wrap rounded px-1 py-0.5 cursor-pointer hover:bg-yellow-50 transition-colors ${
                  isSelected ? "bg-yellow-100 outline outline-1 outline-amber-400" : ""
                }`}
              >
                {paraContent.length > 0 ? paraContent : "\u00a0"}
              </p>
            );
          }

          // Handle drawing blocks (logos, images, etc.)
          if (block.type === "drawing") {
            const drawing = block as DrawingBlock;
            const widthPx = Math.round(drawing.width_inches * 96); // 96 DPI
            const heightPx = Math.round(drawing.height_inches * 96);
            const isDrawingSelected = selectedBlockId === drawing.id;
            
            return (
              <div
                key={drawing.id}
                ref={(el) => {
                  if (el) {
                    previewRefs.current[drawing.id] = el;
                  }
                }}
                onClick={() => {
                  setSelectedBlockId(drawing.id);
                  setSelectedCellId(null);
                  const blockEl = blockListRefs.current[drawing.id];
                  if (blockEl) {
                    blockEl.scrollIntoView({ behavior: "smooth", block: "center" });
                  }
                }}
                className={`flex items-center justify-center rounded border-2 border-dashed p-2 my-2 cursor-pointer transition-colors ${
                  isDrawingSelected 
                    ? "border-amber-400 bg-yellow-100 outline outline-1 outline-amber-400" 
                    : "border-zinc-300 bg-zinc-50 hover:border-amber-400"
                }`}
                style={{
                  width: widthPx > 0 ? `${widthPx}px` : "100px",
                  height: heightPx > 0 ? `${heightPx}px` : "100px",
                  minWidth: "80px",
                  minHeight: "60px",
                }}
              >
                <div className="text-center text-xs text-zinc-500">
                  <div className="text-2xl mb-1">
                    {drawing.drawing_type === "vector_group" ? "🎨" : 
                     drawing.drawing_type === "image" ? "🖼️" : "📄"}
                  </div>
                  <div className="font-medium">{drawing.name || "Drawing"}</div>
                  <div className="text-[10px]">
                    {drawing.width_inches.toFixed(1)}" × {drawing.height_inches.toFixed(1)}"
                  </div>
                  <div className="text-[10px] text-zinc-400">
                    {drawing.drawing_type === "vector_group" ? "Vector Logo" : drawing.drawing_type}
                  </div>
                </div>
              </div>
            );
          }

          const table = block as TableBlock;
          
          // Helper to convert border style to CSS
          const getBorderStyle = (border: CellBorder | null | undefined): string => {
            if (!border || border.style === "none" || border.style === "nil") {
              return "none";
            }
            const width = Math.max(1, Math.round(border.width / 8)); // Convert eighths of pt to px
            const color = border.color ? `#${border.color}` : "#000";
            const style = border.style === "double" ? "double" : 
                         border.style === "dashed" ? "dashed" :
                         border.style === "dotted" ? "dotted" : "solid";
            return `${width}px ${style} ${color}`;
          };
          
          // Recursive helper to render cell content (paragraphs and nested tables)
          const renderCellContent = (cellBlocks: CellBlock[], parentCellId: string): React.ReactNode[] => {
            return cellBlocks.map((cellBlock, idx) => {
              if (cellBlock.type === "paragraph") {
                const p = cellBlock as ParagraphBlock;
                return (
                  <div key={p.id} className="leading-tight">
                    {p.runs.map((r, ri) => {
                      const style: React.CSSProperties = {};
                      if (r.color) {
                        style.color = `#${r.color}`;
                      }
                      if (r.bold) {
                        style.fontWeight = "bold";
                      }
                      if (r.italic) {
                        style.fontStyle = "italic";
                      }
                      return (
                        <span key={`${p.id}-${ri}`} style={style}>
                          {r.text}
                        </span>
                      );
                    })}
                  </div>
                );
              } else if (cellBlock.type === "table") {
                // Nested table - render recursively
                const nestedTable = cellBlock as TableBlock;
                return (
                  <table key={nestedTable.id} className="w-full border-collapse text-xs my-1">
                    <tbody>
                      {nestedTable.rows.map((nestedRow) => (
                        <tr key={nestedRow.id}>
                          {nestedRow.cells.map((nestedCell) => {
                            if (nestedCell.v_merge === "continue") return null;
                            
                            const isNestedSelected = selectedCellId === nestedCell.id;
                            const nestedCellStyle: React.CSSProperties = { padding: "3px" };
                            
                            if (nestedCell.background_color && !isNestedSelected) {
                              nestedCellStyle.backgroundColor = `#${nestedCell.background_color}`;
                            }
                            if (nestedCell.borders) {
                              if (nestedCell.borders.top) nestedCellStyle.borderTop = getBorderStyle(nestedCell.borders.top);
                              if (nestedCell.borders.bottom) nestedCellStyle.borderBottom = getBorderStyle(nestedCell.borders.bottom);
                              if (nestedCell.borders.left) nestedCellStyle.borderLeft = getBorderStyle(nestedCell.borders.left);
                              if (nestedCell.borders.right) nestedCellStyle.borderRight = getBorderStyle(nestedCell.borders.right);
                            } else {
                              nestedCellStyle.border = "1px solid #e4e4e7";
                            }
                            
                            const nestedContent = renderCellContent(nestedCell.blocks, nestedCell.id);
                            
                            return (
                              <td
                                key={nestedCell.id}
                                ref={(el) => { if (el) previewRefs.current[nestedCell.id] = el; }}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setSelectedCellId(nestedCell.id);
                                  setSelectedBlockId(null);
                                }}
                                colSpan={nestedCell.col_span > 1 ? nestedCell.col_span : undefined}
                                rowSpan={nestedCell.row_span > 1 ? nestedCell.row_span : undefined}
                                style={nestedCellStyle}
                                className={`cursor-pointer hover:bg-yellow-50 transition-colors ${
                                  isNestedSelected ? "bg-yellow-100 outline outline-1 outline-amber-400" : ""
                                }`}
                              >
                                {nestedContent.length > 0 ? nestedContent : "\u00a0"}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                );
              }
              return null;
            });
          };
          
          return (
            <table
              key={table.id}
              className="w-full border-collapse text-xs"
            >
              <tbody>
                {table.rows.map((row) => (
                  <tr key={row.id}>
                    {row.cells.map((cell) => {
                      // Skip cells that are continuations of vertical merges
                      if (cell.v_merge === "continue") {
                        return null;
                      }
                      
                      const isSelectedCell = selectedCellId === cell.id;
                      // Build cell content with styled runs and nested tables
                      const cellContent = renderCellContent(cell.blocks, cell.id);
                      
                      // Build cell style with background and borders
                      const cellStyle: React.CSSProperties = {
                        padding: "4px",
                      };
                      
                      if (cell.background_color && !isSelectedCell) {
                        cellStyle.backgroundColor = `#${cell.background_color}`;
                      }
                      
                      // Apply borders
                      if (cell.borders) {
                        if (cell.borders.top) {
                          cellStyle.borderTop = getBorderStyle(cell.borders.top);
                        }
                        if (cell.borders.bottom) {
                          cellStyle.borderBottom = getBorderStyle(cell.borders.bottom);
                        }
                        if (cell.borders.left) {
                          cellStyle.borderLeft = getBorderStyle(cell.borders.left);
                        }
                        if (cell.borders.right) {
                          cellStyle.borderRight = getBorderStyle(cell.borders.right);
                        }
                      } else {
                        // Default border if no specific borders defined
                        cellStyle.border = "1px solid #d4d4d8";
                      }
                      
                      // Calculate rowSpan for vertical merges
                      let rowSpan = cell.row_span;
                      if (cell.v_merge === "restart") {
                        // Count how many continuation cells follow
                        const cellIdx = row.cells.indexOf(cell);
                        const rowIdx = table.rows.indexOf(row);
                        rowSpan = 1;
                        for (let ri = rowIdx + 1; ri < table.rows.length; ri++) {
                          const nextRow = table.rows[ri];
                          const nextCell = nextRow.cells[cellIdx];
                          if (nextCell && nextCell.v_merge === "continue") {
                            rowSpan++;
                          } else {
                            break;
                          }
                        }
                      }
                      
                      return (
                        <td
                          key={cell.id}
                          ref={(el) => {
                            if (el) {
                              previewRefs.current[cell.id] = el;
                            }
                          }}
                          onClick={() => {
                            setSelectedCellId(cell.id);
                            setSelectedBlockId(null);
                            // Scroll to cell in left panel
                            const blockEl = blockListRefs.current[cell.id];
                            if (blockEl) {
                              blockEl.scrollIntoView({ behavior: "smooth", block: "center" });
                            }
                          }}
                          colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                          rowSpan={rowSpan > 1 ? rowSpan : undefined}
                          style={cellStyle}
                          className={`cursor-pointer hover:bg-yellow-50 transition-colors ${
                            isSelectedCell
                              ? "bg-yellow-100 outline outline-1 outline-amber-400"
                              : ""
                          }`}
                        >
                          {cellContent.length > 0 ? cellContent : "\u00a0"}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          );
        })}
      </div>
    );
  };

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;
    await uploadFile(file);
  };

  const selectedBlock: ParagraphBlock | null = (() => {
    if (!doc || !selectedBlockId) return null;
    const block = doc.blocks.find((b) => b.id === selectedBlockId);
    if (!block || block.type !== "paragraph") return null;
    return block as ParagraphBlock;
  })();

  const selectedText = (() => {
    if (selectedBlock) {
      return selectedBlock.runs.map((r) => r.text).join("");
    }
    if (doc && selectedCellId) {
      // Recursive helper to find cell in nested tables
      const findCellText = (blocks: CellBlock[]): string | null => {
        for (const block of blocks) {
          if (block.type === "table") {
            for (const row of block.rows) {
              for (const cell of row.cells) {
                if (cell.id === selectedCellId) {
                  // Get text from paragraph blocks only
                  const paraBlocks = cell.blocks.filter((b): b is ParagraphBlock => b.type === "paragraph");
                  if (paraBlocks.length === 0) return "";
                  return paraBlocks.flatMap((p) => p.runs.map((r) => r.text)).join("");
                }
                // Search nested tables
                const nestedResult = findCellText(cell.blocks);
                if (nestedResult !== null) return nestedResult;
              }
            }
          }
        }
        return null;
      };
      
      const result = findCellText(doc.blocks as CellBlock[]);
      if (result !== null) return result;
    }
    return "";
  })();

  const handleSelectedTextChange = (value: string) => {
    if (!doc) return;

    // Paragraph selected
    if (selectedBlock) {
      const newDoc: DocumentJSON = {
        ...doc,
        blocks: doc.blocks.map((block) => {
          if (block.id !== selectedBlock.id || block.type !== "paragraph") {
            return block;
          }
          const paragraph = block as ParagraphBlock;
          const baseRun = paragraph.runs[0];
          const singleRun: Run = baseRun
            ? { ...baseRun, text: value }
            : {
                id: `${paragraph.id}-run-0`,
                xml_ref: paragraph.xml_ref + "/r[0]",
                text: value,
                bold: false,
                italic: false,
              };
          return { ...paragraph, runs: [singleRun] };
        }),
      };
      setDoc(newDoc);
      return;
    }

    // Cell selected - handle nested tables recursively
    if (selectedCellId) {
      // Helper to update cell text, handling nested tables
      const updateCellInBlocks = (blocks: CellBlock[]): CellBlock[] => {
        return blocks.map((block) => {
          if (block.type !== "table") return block;
          const table = block as TableBlock;
          const newRows = table.rows.map((row) => {
            const newCells = row.cells.map((cell) => {
              if (cell.id === selectedCellId) {
                // Found the cell - update it
                const firstPara = cell.blocks.find((b): b is ParagraphBlock => b.type === "paragraph");
                const para: ParagraphBlock =
                  firstPara ?? {
                    type: "paragraph",
                    id: `${cell.id}-p-0`,
                    xml_ref: `${cell.xml_ref}/p[0]`,
                    style_name: null,
                    runs: [],
                  };
                const baseRun = para.runs[0];
                const singleRun: Run = baseRun
                  ? { ...baseRun, text: value }
                  : {
                      id: `${para.id}-run-0`,
                      xml_ref: `${para.xml_ref}/r[0]`,
                      text: value,
                      bold: false,
                      italic: false,
                    };
                const newPara: ParagraphBlock = { ...para, runs: [singleRun] };
                // Keep nested tables, replace/add paragraph
                const otherBlocks = cell.blocks.filter((b) => b.type !== "paragraph" || b.id !== para.id);
                const newBlocks: CellBlock[] = [newPara, ...otherBlocks.filter((b): b is CellBlock => b.type === "paragraph" || b.type === "table")];
                return { ...cell, blocks: newBlocks };
              }
              // Recursively search nested tables
              const updatedBlocks = updateCellInBlocks(cell.blocks);
              if (updatedBlocks !== cell.blocks) {
                return { ...cell, blocks: updatedBlocks };
              }
              return cell;
            });
            return { ...row, cells: newCells };
          });
          return { ...table, rows: newRows };
        });
      };
      
      const newDoc: DocumentJSON = {
        ...doc,
        blocks: updateCellInBlocks(doc.blocks as CellBlock[]) as Block[],
      };
      setDoc(newDoc);
    }
  };

  const handleToggleCheckbox = async (checkboxId: string, nextChecked: boolean) => {
    if (!doc || !documentId) return;
    
    // Optimistic update
    const newDoc: DocumentJSON = {
      ...doc,
      checkboxes: doc.checkboxes.map((cb) =>
        cb.id === checkboxId ? { ...cb, checked: nextChecked } : cb
      ),
    };
    setDoc(newDoc);
    
    // Persist to backend
    try {
      const res = await fetch(`${BACKEND_URL}/documents/${documentId}/checkbox`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ checkbox_id: checkboxId, checked: nextChecked }),
      });
      if (!res.ok) {
        throw new Error("Checkbox update failed");
      }
      const data = (await res.json()) as DocumentJSON;
      setDoc(normalizeDocument(data));
    } catch (err) {
      console.error(err);
      // Revert on error
      setDoc(doc);
    }
  };

  const handleDropdownChange = async (dropdownId: string, selected: string) => {
    if (!doc || !documentId) return;
    
    // Optimistic update
    const newDoc: DocumentJSON = {
      ...doc,
      dropdowns: doc.dropdowns.map((dd) =>
        dd.id === dropdownId ? { ...dd, selected } : dd
      ),
    };
    setDoc(newDoc);
    
    // Persist to backend
    try {
      const res = await fetch(`${BACKEND_URL}/documents/${documentId}/dropdown`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ dropdown_id: dropdownId, selected }),
      });
      if (!res.ok) {
        throw new Error("Dropdown update failed");
      }
      const data = (await res.json()) as DocumentJSON;
      setDoc(data);
    } catch (err) {
      console.error(err);
      // Revert on error
      setDoc(doc);
    }
  };

  const handleSave = async () => {
    if (!doc || !documentId) return;
    setIsSaving(true);
    try {
      const res = await fetch(`${BACKEND_URL}/documents/${documentId}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(doc),
      });
      if (!res.ok) {
        const detail = await res.json().catch(() => null);
        console.error("Save failed", detail);
        throw new Error("Save failed");
      }
      const data = (await res.json()) as DocumentJSON;
      setDoc(data);
    } catch (err) {
      console.error(err);
      alert("Save failed. Check console for details.");
    } finally {
      setIsSaving(false);
    }
  };

  const [isAIEditing, setIsAIEditing] = useState(false);
  const [isPreviewFullscreen, setIsPreviewFullscreen] = useState(false);

  const handleAIEdit = async () => {
    if (!documentId || (!selectedBlock && !selectedCellId)) return;
    if (!instruction.trim()) {
      alert("Please enter an instruction for the AI edit.");
      return;
    }
    
    setIsAIEditing(true);
    try {
      const res = await fetch(`${BACKEND_URL}/documents/${documentId}/ai-edit`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          block_id: selectedBlock?.id || "",
          cell_id: selectedCellId || null,
          instruction: instruction,
        }),
      });
      if (!res.ok) {
        const detail = await res.json().catch(() => null);
        console.error("AI edit failed", detail);
        throw new Error(detail?.detail || "AI edit failed");
      }
      const data = (await res.json()) as DocumentJSON;
      setDoc(data);
      setInstruction(""); // Clear instruction after successful edit
    } catch (err) {
      console.error(err);
      alert(`AI edit failed: ${err instanceof Error ? err.message : "Unknown error"}`);
    } finally {
      setIsAIEditing(false);
    }
  };

  const handleExport = async () => {
    if (!documentId) return;
    setIsExporting(true);
    try {
      const res = await fetch(
        `${BACKEND_URL}/documents/${documentId}/export/file`,
        {
          method: "POST",
        }
      );
      if (!res.ok) {
        throw new Error(`Export failed: ${res.status}`);
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      const filename = `${documentId.replace(/\.docx$/i, "")}_copy.docx`;
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      setExportPath(filename);
    } catch (err) {
      console.error(err);
      alert("Export failed. Check console for details.");
    } finally {
      setIsExporting(false);
    }
  };

  const handleExcelExport = async () => {
    if (!spreadsheetId) return;
    setIsExporting(true);
    try {
      const res = await fetch(
        `${BACKEND_URL}/spreadsheets/${spreadsheetId}/export/file`,
        {
          method: "POST",
        }
      );
      if (!res.ok) {
        throw new Error(`Export failed: ${res.status}`);
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      const filename = `${spreadsheetId.replace(/\.xlsx$/i, "")}_copy.xlsx`;
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      setExportPath(filename);
    } catch (err) {
      console.error(err);
      alert("Export failed. Check console for details.");
    } finally {
      setIsExporting(false);
    }
  };

  // Get selected Excel cell data
  const getSelectedExcelCellData = (): ExcelCell | null => {
    if (!spreadsheet || !selectedExcelCell) return null;
    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) return null;
    return activeSheet.cells.find(c => c.ref === selectedExcelCell) || null;
  };

  // Handle Excel cell selection
  const handleExcelCellSelect = (ref: string) => {
    setSelectedExcelCell(ref);
    const activeSheet = spreadsheet?.sheets[activeSheetIndex];
    if (activeSheet) {
      const cell = activeSheet.cells.find(c => c.ref === ref);
      setExcelCellEditValue(cell?.value !== null && cell?.value !== undefined ? String(cell.value) : "");
    }
  };

  // Save Excel cell edit
  const handleSaveExcelCell = async () => {
    if (!spreadsheetId || !selectedExcelCell || !spreadsheet) return;
    
    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) return;
    
    setIsSavingExcelCell(true);
    try {
      const res = await fetch(`${BACKEND_URL}/spreadsheets/${spreadsheetId}/cell`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheet: activeSheet.name,
          cell: selectedExcelCell,
          value: excelCellEditValue,
        }),
      });
      
      if (!res.ok) {
        throw new Error(`Save failed: ${res.status}`);
      }
      
      const data = (await res.json()) as SpreadsheetJSON;
      setSpreadsheet(data);
    } catch (err) {
      console.error(err);
      alert("Save failed. Check console for details.");
    } finally {
      setIsSavingExcelCell(false);
    }
  };

  // Excel AI Edit
  const handleExcelAIEdit = async () => {
    if (!spreadsheetId || !selectedExcelCell || !spreadsheet) return;
    if (!instruction.trim()) {
      alert("Please enter an instruction for the AI edit.");
      return;
    }
    
    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) return;
    
    setIsAIEditing(true);
    try {
      const res = await fetch(`${BACKEND_URL}/spreadsheets/${spreadsheetId}/ai-edit`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheet: activeSheet.name,
          cell: selectedExcelCell,
          instruction: instruction,
        }),
      });
      
      if (!res.ok) {
        const detail = await res.json().catch(() => null);
        console.error("AI edit failed", detail);
        throw new Error(detail?.detail || "AI edit failed");
      }
      
      const data = (await res.json()) as SpreadsheetJSON;
      setSpreadsheet(data);
      setInstruction(""); // Clear instruction after successful edit
      
      // Update the cell edit value to reflect the change
      const updatedSheet = data.sheets[activeSheetIndex];
      if (updatedSheet) {
        const updatedCell = updatedSheet.cells.find(c => c.ref === selectedExcelCell);
        if (updatedCell) {
          setExcelCellEditValue(updatedCell.value !== null && updatedCell.value !== undefined ? String(updatedCell.value) : "");
        }
      }
    } catch (err) {
      console.error(err);
      alert(`AI edit failed: ${err instanceof Error ? err.message : "Unknown error"}`);
    } finally {
      setIsAIEditing(false);
    }
  };

  // Excel Checkbox Toggle
  const handleExcelCheckboxToggle = async (sheetName: string, controlId: string, checked: boolean) => {
    if (!spreadsheetId || !spreadsheet) return;
    
    try {
      const res = await fetch(`${BACKEND_URL}/spreadsheets/${spreadsheetId}/checkbox`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheet: sheetName,
          control_id: controlId,
          checked: checked,
        }),
      });
      
      if (!res.ok) {
        throw new Error("Checkbox update failed");
      }
      
      const data = (await res.json()) as SpreadsheetJSON;
      setSpreadsheet(data);
    } catch (err) {
      console.error(err);
      alert("Checkbox update failed. Check console for details.");
    }
  };

  // Excel Dropdown Change
  const handleExcelDropdownChange = async (sheetName: string, cellRef: string, value: string) => {
    if (!spreadsheetId || !spreadsheet) return;
    
    try {
      const res = await fetch(`${BACKEND_URL}/spreadsheets/${spreadsheetId}/dropdown`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheet: sheetName,
          cell: cellRef,
          value: value,
        }),
      });
      
      if (!res.ok) {
        const detail = await res.json().catch(() => null);
        throw new Error(detail?.detail || "Dropdown update failed");
      }
      
      const data = (await res.json()) as SpreadsheetJSON;
      setSpreadsheet(data);
    } catch (err) {
      console.error(err);
      alert(`Dropdown update failed: ${err instanceof Error ? err.message : "Unknown error"}`);
    }
  };

  // Get dropdowns for current sheet
  const getExcelDropdowns = (): { cellRef: string; options: string[]; currentValue: string | null }[] => {
    if (!spreadsheet) return [];
    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) return [];
    
    const dropdowns: { cellRef: string; options: string[]; currentValue: string | null }[] = [];
    
    for (const cell of activeSheet.cells) {
      if (cell.dropdown && cell.dropdown.options.length > 0) {
        dropdowns.push({
          cellRef: cell.ref,
          options: cell.dropdown.options,
          currentValue: cell.value !== null && cell.value !== undefined ? String(cell.value) : null,
        });
      }
    }
    
    return dropdowns;
  };

  // Get checkboxes for current sheet
  const getExcelCheckboxes = (): { controlId: string; checked: boolean; linkedCell: string | null }[] => {
    if (!spreadsheet) return [];
    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) return [];
    
    return activeSheet.form_controls
      .filter(c => c.type === "checkbox")
      .map(c => ({
        controlId: c.id,
        checked: c.checked ?? false,
        linkedCell: c.linked_cell ?? null,
      }));
  };

  // =============================================================================
  // EXCEL PREVIEW RENDERER
  // =============================================================================
  
  const renderExcelPreview = () => {
    if (!spreadsheet) {
      return (
        <p className="text-sm text-zinc-500">
          Upload an Excel file to see the preview.
        </p>
      );
    }

    const activeSheet = spreadsheet.sheets[activeSheetIndex];
    if (!activeSheet) {
      return <p className="text-sm text-zinc-500">No sheet selected.</p>;
    }

    // Build grid from cells
    const cells = activeSheet.cells;
    if (cells.length === 0) {
      return <p className="text-sm text-zinc-500">Empty sheet.</p>;
    }

    // Find grid bounds
    const maxRow = Math.min(100, Math.max(...cells.map(c => c.row))); // Limit to 100 rows
    const maxCol = Math.min(26, Math.max(...cells.map(c => c.col))); // Limit to 26 cols (A-Z)

    // Create cell lookup
    const cellMap: Record<string, ExcelCell> = {};
    for (const cell of cells) {
      cellMap[cell.ref] = cell;
    }

    // Generate column headers (A, B, C, ...)
    const colLetters = Array.from({ length: maxCol }, (_, i) => 
      String.fromCharCode(65 + i)
    );

    return (
      <div className="flex flex-col h-full overflow-hidden">
        {/* Sheet tabs - horizontal scroll */}
        <div className="flex-shrink-0 overflow-x-auto border-b pb-2 mb-2">
          <div className="flex gap-1 min-w-max">
            {spreadsheet.sheets.map((sheet, idx) => (
              <button
                key={sheet.id}
                onClick={() => {
                  setActiveSheetIndex(idx);
                  setSelectedExcelCell(null); // Clear selection when switching sheets
                  setExcelCellEditValue("");
                }}
                className={`px-3 py-1 text-xs rounded-t whitespace-nowrap ${
                  idx === activeSheetIndex
                    ? "bg-emerald-100 text-emerald-800 font-medium"
                    : "bg-zinc-100 text-zinc-600 hover:bg-zinc-200"
                }`}
              >
                {sheet.name}
                {sheet.cell_count > 5000 && (
                  <span className="ml-1 text-[10px] text-zinc-400">
                    ({sheet.cell_count.toLocaleString()})
                  </span>
                )}
              </button>
            ))}
          </div>
        </div>

        {/* Sheet stats */}
        <div className="flex-shrink-0 flex flex-wrap gap-3 mb-2 text-[10px] text-zinc-500">
          <span>Cells: {activeSheet.cell_count.toLocaleString()}</span>
          <span>Merged: {activeSheet.merged_cells.length}</span>
          <span>Validations: {activeSheet.data_validations.length}</span>
          <span>CF Rules: {activeSheet.conditional_formatting_count}</span>
          {activeSheet.freeze_pane && (
            <span className="text-emerald-600">
              Frozen: {activeSheet.freeze_pane.rows}R/{activeSheet.freeze_pane.cols}C
            </span>
          )}
        </div>

        {/* Grid - scrollable container */}
        <div className="flex-1 overflow-auto min-h-0">
          <table className="border-collapse text-xs">
          <thead>
            <tr>
              <th className="w-8 bg-zinc-100 border border-zinc-300 text-zinc-500"></th>
              {colLetters.map(letter => (
                <th key={letter} className="px-2 py-1 bg-zinc-100 border border-zinc-300 text-zinc-600 min-w-[80px]">
                  {letter}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: maxRow }, (_, rowIdx) => {
              const rowNum = rowIdx + 1;
              return (
                <tr key={rowNum}>
                  <td className="px-2 py-1 bg-zinc-100 border border-zinc-300 text-zinc-500 text-center">
                    {rowNum}
                  </td>
                  {colLetters.map(letter => {
                    const ref = `${letter}${rowNum}`;
                    const cell = cellMap[ref];
                    const isSelected = selectedExcelCell === ref;
                    
                    // Skip merged continuation cells
                    if (cell && cell.is_merged && !cell.is_merge_origin) {
                      return null;
                    }

                    // Calculate spans for merge origins
                    let colSpan = 1;
                    let rowSpan = 1;
                    if (cell?.is_merge_origin && cell.merge_range) {
                      const merge = activeSheet.merged_cells.find(m => m.ref === cell.merge_range);
                      if (merge) {
                        colSpan = merge.end_col - merge.start_col + 1;
                        rowSpan = merge.end_row - merge.start_row + 1;
                      }
                    }

                    const cellStyle: React.CSSProperties = {};
                    if (cell?.style?.bg_color) {
                      // Excel stores colors without # prefix
                      const bgColor = cell.style.bg_color.startsWith('#') 
                        ? cell.style.bg_color 
                        : `#${cell.style.bg_color}`;
                      cellStyle.backgroundColor = bgColor;
                    }
                    if (cell?.style?.color) {
                      const textColor = cell.style.color.startsWith('#') 
                        ? cell.style.color 
                        : `#${cell.style.color}`;
                      cellStyle.color = textColor;
                    }
                    if (cell?.style?.bold) {
                      cellStyle.fontWeight = "bold";
                    }
                    if (cell?.style?.italic) {
                      cellStyle.fontStyle = "italic";
                    }
                    if (cell?.style?.underline) {
                      cellStyle.textDecoration = "underline";
                    }
                    if (cell?.style?.font_size) {
                      cellStyle.fontSize = `${cell.style.font_size}pt`;
                    }
                    // Text alignment
                    if (cell?.style?.h_align) {
                      cellStyle.textAlign = cell.style.h_align as React.CSSProperties["textAlign"];
                    }
                    if (cell?.style?.v_align) {
                      cellStyle.verticalAlign = cell.style.v_align === "center" ? "middle" : cell.style.v_align as React.CSSProperties["verticalAlign"];
                    }
                    // Borders
                    const borderMap: Record<string, string> = {
                      thin: "1px solid #333",
                      medium: "2px solid #333",
                      thick: "3px solid #333",
                      dashed: "1px dashed #333",
                      dotted: "1px dotted #333",
                    };
                    if (cell?.style?.borders) {
                      if (cell.style.borders.left) cellStyle.borderLeft = borderMap[cell.style.borders.left] || "1px solid #333";
                      if (cell.style.borders.right) cellStyle.borderRight = borderMap[cell.style.borders.right] || "1px solid #333";
                      if (cell.style.borders.top) cellStyle.borderTop = borderMap[cell.style.borders.top] || "1px solid #333";
                      if (cell.style.borders.bottom) cellStyle.borderBottom = borderMap[cell.style.borders.bottom] || "1px solid #333";
                    }

                    return (
                      <td
                        key={ref}
                        colSpan={colSpan > 1 ? colSpan : undefined}
                        rowSpan={rowSpan > 1 ? rowSpan : undefined}
                        onClick={() => handleExcelCellSelect(ref)}
                        style={cellStyle}
                        className={`px-2 py-1 border border-zinc-200 cursor-pointer hover:bg-yellow-50/50 transition-colors ${
                          isSelected ? "bg-yellow-100 outline outline-2 outline-emerald-400" : ""
                        } ${cell?.formula ? "text-blue-600" : ""} ${cell?.style?.wrap ? "whitespace-pre-wrap" : ""}`}
                      >
                        {cell?.value !== null && cell?.value !== undefined 
                          ? String(cell.value) 
                          : ""}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
        </div>{/* End grid scrollable container */}
      </div>
    );
  };

  return (
    <div className="h-screen bg-zinc-50 text-zinc-900">
      <div className="flex h-screen w-full flex-col gap-4 px-6 py-4">
        <header className="flex items-center justify-between border-b pb-2">
          <div className="flex items-center gap-6">
            <div>
              <h1 className="text-xl font-semibold">Document Digital Copy POC</h1>
              <p className="text-sm text-zinc-600">
                Upload, view, edit, and export documents with full structural fidelity.
              </p>
            </div>
            {/* Mode Toggle Pills */}
            <div className="flex gap-2 rounded-full bg-zinc-100 p-1">
              <ModePill 
                mode="docx" 
                currentMode={mode} 
                onClick={() => setMode("docx")} 
                label="📄 DOCX" 
              />
              <ModePill 
                mode="excel" 
                currentMode={mode} 
                onClick={() => setMode("excel")} 
                label="📊 Excel" 
              />
            </div>
          </div>
          {mode === "docx" && documentId && (
            <span className="rounded bg-zinc-200 px-2 py-1 text-xs font-mono">
              {documentId}
            </span>
          )}
          {mode === "excel" && spreadsheetId && (
            <span className="rounded bg-emerald-100 px-2 py-1 text-xs font-mono">
              {spreadsheetId}
            </span>
          )}
        </header>

        <section className="rounded-md border bg-white p-3 shadow-sm">
          <div className="flex flex-wrap items-center gap-4">
            <input
              ref={fileInputRef}
              type="file"
              name="file"
              accept={mode === "excel" ? ".xlsx,.xls" : ".docx"}
              className="hidden"
              onChange={handleFileChange}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className={`rounded border px-3 py-1.5 text-sm font-medium ${
                mode === "excel" 
                  ? "border-emerald-300 text-emerald-700 hover:bg-emerald-50" 
                  : "text-zinc-800"
              }`}
            >
              {mode === "excel" ? "Choose Excel File" : "Choose DOCX File"}
            </button>
            <button
              type="button"
              disabled={isUploading}
              className={`rounded px-3 py-1.5 text-sm font-medium text-white disabled:opacity-60 ${
                mode === "excel" ? "bg-emerald-600" : "bg-zinc-900"
              }`}
            >
              {isUploading ? "Uploading..." : "Upload on Select"}
            </button>
            
            {/* DOCX Export */}
            {mode === "docx" && documentId && (
              <button
                type="button"
                onClick={handleExport}
                disabled={isExporting}
                className="rounded border px-3 py-1.5 text-sm font-medium text-zinc-800 disabled:opacity-60"
              >
                {isExporting ? "Exporting..." : "Export DOCX"}
              </button>
            )}
            
            {/* Excel Export */}
            {mode === "excel" && spreadsheetId && (
              <button
                type="button"
                onClick={handleExcelExport}
                disabled={isExporting}
                className="rounded border border-emerald-300 px-3 py-1.5 text-sm font-medium text-emerald-700 disabled:opacity-60"
              >
                {isExporting ? "Exporting..." : "Export Excel"}
              </button>
            )}
            
            {/* Fullscreen toggle */}
            {(doc || spreadsheet) && (
              <button
                type="button"
                onClick={() => setIsPreviewFullscreen((v) => !v)}
                className="ml-auto rounded border px-3 py-1.5 text-sm font-medium text-zinc-800"
              >
                {isPreviewFullscreen ? "Exit Fullscreen" : "Fullscreen Preview"}
              </button>
            )}
          </div>
          {exportPath && (
            <p className="mt-2 text-xs text-zinc-600">
              Latest export: <code>{exportPath}</code>
            </p>
          )}
        </section>

        <main className="flex flex-1 min-h-0 gap-3">
          {!isPreviewFullscreen && (
            <>
          {/* Left: Blocks / Cells */}
          <section className="flex h-full min-h-0 w-1/5 flex-col rounded-md border bg-white p-3 shadow-sm">
            <h2 className={`mb-2 text-sm font-semibold ${mode === "excel" ? "text-emerald-700" : ""}`}>
              {mode === "excel" ? "📊 Cells" : "Blocks"}
            </h2>
            <div className="flex-1 space-y-1 overflow-auto text-sm">
              {/* Excel mode: show cells */}
              {mode === "excel" && spreadsheet ? (
                (() => {
                  const activeSheet = spreadsheet.sheets[activeSheetIndex];
                  if (!activeSheet) return <p className="text-xs text-zinc-500">No sheet.</p>;
                  
                  // Group cells with values
                  const cellsWithValues = activeSheet.cells.filter(
                    c => c.value !== null && c.value !== undefined && String(c.value).trim() !== ""
                  ).slice(0, 200); // Limit to 200 for performance
                  
                  if (cellsWithValues.length === 0) {
                    return <p className="text-xs text-zinc-500">No cells with data.</p>;
                  }
                  
                  return (
                    <>
                      <div className="text-[10px] text-zinc-400 mb-2">
                        {activeSheet.name} • {cellsWithValues.length} cells shown
                      </div>
                      {cellsWithValues.map((cell) => (
                        <button
                          key={cell.id}
                          onClick={() => handleExcelCellSelect(cell.ref)}
                          className={`block w-full rounded px-2 py-1 text-left hover:bg-emerald-50 transition-colors ${
                            selectedExcelCell === cell.ref ? "bg-emerald-100" : ""
                          }`}
                        >
                          <span className="mr-2 text-xs font-mono text-emerald-600">
                            {cell.ref}
                          </span>
                          <span className="truncate text-xs text-zinc-700">
                            {String(cell.value).substring(0, 40)}
                            {String(cell.value).length > 40 ? "..." : ""}
                          </span>
                          {cell.formula && (
                            <span className="ml-1 text-[10px] text-blue-500">fx</span>
                          )}
                        </button>
                      ))}
                    </>
                  );
                })()
              ) : mode === "excel" ? (
                <p className="text-sm text-zinc-500">
                  Upload an Excel file to see cells.
                </p>
              ) : doc ? (
                doc.blocks.map((block, index) => {
                  if (block.type === "paragraph") {
                    const para = block as ParagraphBlock;
                    const text = para.runs.map((r) => r.text).join("");
                    return (
                      <button
                        key={para.id}
                        ref={(el) => {
                          if (el) {
                            blockListRefs.current[para.id] = el;
                          }
                        }}
                        className={`block w-full rounded px-2 py-1 text-left hover:bg-zinc-100 transition-colors ${
                          selectedBlockId === para.id ? "bg-zinc-200" : ""
                        }`}
                        onClick={() => {
                          setSelectedBlockId(para.id);
                          setSelectedCellId(null);
                          const el = previewRefs.current[para.id];
                          if (el) {
                            el.scrollIntoView({ behavior: "smooth", block: "center" });
                          }
                        }}
                      >
                        <span className="mr-1 text-xs font-mono text-zinc-500">
                          P{index}
                        </span>
                        <span className="truncate text-xs">
                          {text || "(empty paragraph)"}
                        </span>
                      </button>
                    );
                  }

                  if (block.type === "table") {
                    const table = block as TableBlock;
                    
                    // Helper to get cell text (handles nested tables)
                    const getCellText = (cellBlocks: CellBlock[]): string => {
                      return cellBlocks
                        .filter((b): b is ParagraphBlock => b.type === "paragraph")
                        .flatMap((p) => p.runs.map((r) => r.text))
                        .join(" ");
                    };
                    
                    // Recursive helper to render table cells in block list
                    const renderTableCells = (
                      tbl: TableBlock,
                      prefix: string,
                      depth: number
                    ): React.ReactNode[] => {
                      const nodes: React.ReactNode[] = [];
                      tbl.rows.forEach((row, rIdx) => {
                        row.cells.forEach((cell, cIdx) => {
                          const cellText = getCellText(cell.blocks);
                          const hasNested = cell.blocks.some((b) => b.type === "table");
                          
                          nodes.push(
                            <button
                              key={cell.id}
                              ref={(el) => { if (el) blockListRefs.current[cell.id] = el; }}
                              className={`block w-full rounded px-3 py-0.5 text-left hover:bg-zinc-100 transition-colors ${
                                selectedCellId === cell.id ? "bg-zinc-200" : ""
                              }`}
                              style={{ paddingLeft: `${12 + depth * 8}px` }}
                              onClick={() => {
                                setSelectedCellId(cell.id);
                                setSelectedBlockId(null);
                                const el = previewRefs.current[cell.id];
                                if (el) {
                                  el.scrollIntoView({ behavior: "smooth", block: "center" });
                                }
                              }}
                            >
                              <span className="mr-1 font-mono text-[10px] text-zinc-500">
                                {prefix}R{rIdx}C{cIdx}
                              </span>
                              <span className="truncate text-[11px]">
                                {cellText || (hasNested ? "(nested table)" : "(empty cell)")}
                              </span>
                            </button>
                          );
                          
                          // Recursively render nested tables
                          cell.blocks
                            .filter((b): b is TableBlock => b.type === "table")
                            .forEach((nestedTbl, ntIdx) => {
                              nodes.push(
                                <div key={nestedTbl.id} className="text-[10px] text-zinc-500 px-3 py-0.5" style={{ paddingLeft: `${16 + depth * 8}px` }}>
                                  ↳ Nested table ({nestedTbl.rows.length} rows)
                                </div>
                              );
                              nodes.push(...renderTableCells(nestedTbl, `${prefix}N${ntIdx}-`, depth + 1));
                            });
                        });
                      });
                      return nodes;
                    };
                    
                    return (
                      <div key={table.id} className="space-y-0.5 text-xs">
                        <div className="rounded px-2 py-1 text-zinc-700">
                          <span className="mr-1 font-mono text-zinc-500">
                            T{index}
                          </span>
                          Table with {table.rows.length} row(s)
                        </div>
                        {renderTableCells(table, "", 0)}
                      </div>
                    );
                  }

                  // Drawing block: show a simple entry
                  const drawing = block as DrawingBlock;
                  const isDrawingSelected = selectedBlockId === drawing.id;
                  return (
                    <button
                      key={drawing.id}
                      ref={(el) => {
                        if (el) {
                          blockListRefs.current[drawing.id] = el;
                        }
                      }}
                      className={`block w-full rounded px-2 py-1 text-xs text-left flex items-center gap-2 hover:bg-zinc-100 transition-colors ${
                        isDrawingSelected ? "bg-zinc-200" : "text-zinc-600"
                      }`}
                      onClick={() => {
                        setSelectedBlockId(drawing.id);
                        setSelectedCellId(null);
                        const el = previewRefs.current[drawing.id];
                        if (el) {
                          el.scrollIntoView({ behavior: "smooth", block: "center" });
                        }
                      }}
                    >
                      <span className="font-mono text-zinc-500">D{index}</span>
                      <span className="truncate">
                        {drawing.name || "Drawing"} ({drawing.drawing_type})
                      </span>
                    </button>
                  );
                })
              ) : (
                <p className="text-sm text-zinc-500">
                  No document loaded.
                </p>
              )}
            </div>
          </section>

          {/* Middle: Editor */}
          <section className="flex h-full min-h-0 w-2/5 flex-col rounded-md border bg-white shadow-sm overflow-hidden">
            <div className="flex items-center justify-between gap-2 border-b px-4 py-3 flex-shrink-0">
              <h2 className={`text-sm font-semibold ${mode === "excel" ? "text-emerald-700" : ""}`}>
                {mode === "excel" ? "📊 Cell Editor" : "Editor"}
              </h2>
              {mode === "excel" && selectedExcelCell ? (
                <div className="text-[11px] text-emerald-600">
                  Cell <code className="font-mono font-bold">{selectedExcelCell}</code>
                  {getSelectedExcelCellData()?.formula && (
                    <span className="ml-2 text-blue-500">has formula</span>
                  )}
                </div>
              ) : (selectedBlock || selectedCellId) ? (
                <div className="text-[11px] text-zinc-500">
                  {selectedBlock && (
                    <span>
                      Paragraph <code className="font-mono">{selectedBlock.id}</code>
                    </span>
                  )}
                  {selectedCellId && !selectedBlock && (
                    <span>
                      Table cell <code className="font-mono">{selectedCellId}</code>
                    </span>
                  )}
                </div>
              ) : null}
            </div>

            {/* Scrollable content area */}
            <div className="flex-1 overflow-auto px-4 py-3 space-y-4">
              {/* Excel cell editor */}
              {mode === "excel" ? (
                <div className="flex flex-col gap-3">
                  <label className="text-xs font-medium text-emerald-700">
                    Cell Value
                  </label>
                  {selectedExcelCell ? (
                    <>
                      <textarea
                        className="min-h-[7rem] w-full resize-vertical rounded border border-emerald-200 px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/60"
                        value={excelCellEditValue}
                        onChange={(e) => setExcelCellEditValue(e.target.value)}
                        placeholder="Enter cell value..."
                      />
                      {getSelectedExcelCellData()?.formula && (
                        <div className="text-xs text-blue-600 bg-blue-50 px-2 py-1 rounded">
                          <span className="font-medium">Formula:</span> {getSelectedExcelCellData()?.formula}
                        </div>
                      )}
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          type="button"
                          onClick={handleSaveExcelCell}
                          disabled={isSavingExcelCell}
                          className="rounded bg-emerald-600 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-60"
                        >
                          {isSavingExcelCell ? "Saving..." : "Save Cell"}
                        </button>
                        <span className="text-[10px] text-zinc-400">
                          Changes are saved to the spreadsheet JSON
                        </span>
                      </div>
                    </>
                  ) : (
                    <p className="text-sm text-zinc-500">
                      Click a cell in the preview or select from the Cells list to edit.
                    </p>
                  )}

                  {/* Excel AI Edit Section */}
                  <div className="space-y-2 rounded-md bg-emerald-50 px-3 py-3">
                    <label
                      htmlFor="excel-ai-instruction"
                      className="text-xs font-medium text-emerald-700"
                    >
                      AI Instruction
                    </label>
                    <textarea
                      id="excel-ai-instruction"
                      className="h-20 w-full resize-vertical rounded border border-emerald-200 px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-emerald-500/60"
                      value={instruction}
                      onChange={(e) => setInstruction(e.target.value)}
                      placeholder="e.g. make more formal, fix grammar, translate to Spanish..."
                    />
                    <div className="flex flex-wrap gap-2">
                      <button
                        type="button"
                        onClick={handleExcelAIEdit}
                        disabled={isAIEditing || !instruction.trim() || !selectedExcelCell}
                        className="rounded bg-blue-600 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-40"
                      >
                        {isAIEditing ? "Processing..." : "Run AI Edit"}
                      </button>
                      <p className="text-[11px] text-emerald-600">
                        Select a cell with text and enter an instruction.
                      </p>
                    </div>
                  </div>

                  {/* Excel Checkboxes */}
                  {spreadsheet && getExcelCheckboxes().length > 0 && (
                    <div className="space-y-2 border-t border-emerald-200 pt-3">
                      <h3 className="text-xs font-semibold text-emerald-700">
                        Checkboxes (Form Controls)
                      </h3>
                      <div className="space-y-1 max-h-40 overflow-auto pr-1">
                        {getExcelCheckboxes().map((cb) => (
                          <label
                            key={cb.controlId}
                            className="flex cursor-pointer items-start gap-2 text-[11px] text-emerald-700"
                          >
                            <input
                              type="checkbox"
                              className="mt-[2px] h-3 w-3 accent-emerald-600"
                              checked={cb.checked}
                              onChange={(e) =>
                                handleExcelCheckboxToggle(
                                  spreadsheet.sheets[activeSheetIndex]?.name || "",
                                  cb.controlId,
                                  e.target.checked
                                )
                              }
                            />
                            <span className="leading-snug">
                              {cb.linkedCell ? `Linked to ${cb.linkedCell}` : cb.controlId}
                            </span>
                          </label>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Excel Dropdowns (Data Validations) */}
                  {spreadsheet && getExcelDropdowns().length > 0 && (
                    <div className="space-y-2 border-t border-emerald-200 pt-3">
                      <h3 className="text-xs font-semibold text-emerald-700">
                        Dropdowns (Data Validations)
                      </h3>
                      <div className="space-y-2 max-h-48 overflow-auto pr-1">
                        {getExcelDropdowns().map((dd) => (
                          <div key={dd.cellRef} className="flex flex-col gap-1">
                            <label
                              className="text-[11px] text-emerald-600"
                              htmlFor={`excel-dropdown-${dd.cellRef}`}
                            >
                              Cell {dd.cellRef}
                            </label>
                            <select
                              id={`excel-dropdown-${dd.cellRef}`}
                              className="rounded border border-emerald-200 px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-emerald-500/60"
                              value={dd.currentValue || ""}
                              onChange={(e) =>
                                handleExcelDropdownChange(
                                  spreadsheet.sheets[activeSheetIndex]?.name || "",
                                  dd.cellRef,
                                  e.target.value
                                )
                              }
                            >
                              <option value="">-- Select --</option>
                              {dd.options.map((opt) => (
                                <option key={opt} value={opt}>
                                  {opt}
                                </option>
                              ))}
                            </select>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  
                  {/* Excel info panel */}
                  {spreadsheet && (
                    <div className="mt-4 p-3 bg-emerald-50 rounded-md text-xs space-y-2">
                      <h4 className="font-semibold text-emerald-800">Spreadsheet Info</h4>
                      <div className="grid grid-cols-2 gap-1 text-emerald-700">
                        <span>Sheets:</span>
                        <span>{spreadsheet.sheets.length}</span>
                        <span>Active:</span>
                        <span>{spreadsheet.sheets[activeSheetIndex]?.name}</span>
                        <span>Defined Names:</span>
                        <span>{spreadsheet.defined_names.length}</span>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
              /* DOCX text editor section */
              <div className="flex flex-col gap-2">
                <label
                  htmlFor="editor-textarea"
                  className="text-xs font-medium text-zinc-700"
                >
                  Selected text
                </label>
                {selectedBlock || selectedCellId ? (
                  <>
                    <textarea
                      id="editor-textarea"
                      className="min-h-[7rem] w-full resize-vertical rounded border px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/60"
                      value={selectedText}
                      onChange={(e) => handleSelectedTextChange(e.target.value)}
                    />
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={handleSave}
                        disabled={isSaving}
                        className="rounded bg-zinc-900 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-60"
                      >
                        {isSaving ? "Saving..." : "Save JSON"}
                      </button>
                    </div>
                  </>
                ) : (
                  <p className="text-sm text-zinc-500">
                    Select a paragraph (P#) or table cell (R#C#) from the left column
                    to edit its text.
                  </p>
                )}
              </div>
              )}

              {/* AI section - DOCX only */}
              {mode === "docx" && (
              <div className="space-y-2 rounded-md bg-zinc-50 px-3 py-3">
              <label
                htmlFor="ai-instruction"
                className="text-xs font-medium text-zinc-700"
              >
                AI instruction
              </label>
              <textarea
                id="ai-instruction"
                className="h-20 w-full resize-vertical rounded border px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/60"
                value={instruction}
                onChange={(e) => setInstruction(e.target.value)}
                placeholder="e.g. make more formal, fix grammar, make concise..."
              />
              <div className="flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleAIEdit}
                  disabled={isAIEditing || !instruction.trim() || (!selectedBlock && !selectedCellId)}
                  className="rounded bg-blue-600 px-3 py-1.5 text-xs font-medium text-white disabled:opacity-40"
                >
                  {isAIEditing ? "Processing..." : "Run AI Edit"}
                </button>
                <p className="text-[11px] text-zinc-500">
                  Requires a selected block or cell and a non-empty instruction.
                </p>
              </div>
            </div>
              )}

              {/* OOXML Checkboxes - DOCX only */}
              {mode === "docx" && doc && (doc.checkboxes?.length ?? 0) > 0 && (
                <div className="space-y-2 border-t pt-3">
                <h3 className="text-xs font-semibold text-zinc-700">
                  Checkboxes (OOXML)
                </h3>
                <div className="space-y-1 max-h-40 overflow-auto pr-1">
                  {doc.checkboxes.map((cb: CheckboxField) => (
                    <label
                      key={cb.id}
                      className="flex cursor-pointer items-start gap-2 text-[11px] text-zinc-700"
                    >
                      <input
                        type="checkbox"
                        className="mt-[2px] h-3 w-3"
                        checked={cb.checked}
                        onChange={(e) =>
                          handleToggleCheckbox(cb.id, e.target.checked)
                        }
                      />
                      <span className="leading-snug">{cb.label || cb.id}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}

              {/* OOXML Dropdowns - DOCX only */}
              {mode === "docx" && doc && (doc.dropdowns?.length ?? 0) > 0 && (
                <div className="space-y-2 border-t pt-3">
                <h3 className="text-xs font-semibold text-zinc-700">
                  Dropdowns (OOXML)
                </h3>
                <div className="space-y-2 max-h-40 overflow-auto pr-1">
                  {doc.dropdowns.map((dd: DropdownField) => (
                    <div key={dd.id} className="flex flex-col gap-1">
                      <label
                        className="text-[11px] text-zinc-600"
                        htmlFor={`dropdown-${dd.id}`}
                      >
                        {dd.label || dd.id}
                      </label>
                      <select
                        id={`dropdown-${dd.id}`}
                        className="rounded border px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/60"
                        value={dd.selected || ""}
                        onChange={(e) => handleDropdownChange(dd.id, e.target.value)}
                      >
                        {dd.options.map((opt) => (
                          <option key={opt} value={opt}>
                            {opt}
                          </option>
                        ))}
                      </select>
                    </div>
                  ))}
                </div>
              </div>
            )}
            </div>{/* End scrollable content area */}
          </section>
          </>
          )}

          {/* Right: Preview */}
          <section className="flex h-full min-h-0 flex-1 flex-col rounded-md border bg-white p-4 shadow-sm overflow-hidden">
            <h2 className={`mb-2 text-sm font-semibold flex-shrink-0 ${mode === "excel" ? "text-emerald-700" : ""}`}>
              {mode === "excel" ? "📊 Spreadsheet Preview" : "📄 Document Preview"}
            </h2>
            <div className="flex-1 min-h-0 overflow-hidden">
              {mode === "excel" ? renderExcelPreview() : renderPreview()}
            </div>
          </section>
        </main>
      </div>
    </div>
  );
}
