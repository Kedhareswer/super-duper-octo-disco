"""Debug script to identify where text is being lost."""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from services.document_engine import docx_to_json, apply_json_to_docx
from models.schemas import ParagraphBlock, TableBlock

TEST_FILE = Path("data/uploads/docx/test2.docx")
EXPORT_FILE = Path("data/test_outputs/test2_unmodified.docx")


def extract_all_text_detailed(doc, label):
    """Extract all text with location info."""
    texts = []
    
    for bi, block in enumerate(doc.blocks):
        if isinstance(block, ParagraphBlock):
            for ri, run in enumerate(block.runs):
                if run.text:
                    texts.append({
                        'location': f"block[{bi}]/para/run[{ri}]",
                        'xml_ref': run.xml_ref,
                        'text': run.text[:50],
                        'full_len': len(run.text),
                    })
        elif isinstance(block, TableBlock):
            for row_i, row in enumerate(block.rows):
                for cell_i, cell in enumerate(row.cells):
                    for para_i, para in enumerate(cell.blocks):
                        for run_i, run in enumerate(para.runs):
                            if run.text:
                                texts.append({
                                    'location': f"block[{bi}]/tbl/row[{row_i}]/cell[{cell_i}]/para[{para_i}]/run[{run_i}]",
                                    'xml_ref': run.xml_ref,
                                    'text': run.text[:50],
                                    'full_len': len(run.text),
                                })
    
    print(f"\n{'='*70}")
    print(f"  {label}: {len(texts)} text segments")
    print(f"{'='*70}")
    
    return texts


def compare_texts(orig_texts, new_texts):
    """Compare text lists and find differences."""
    print(f"\n{'='*70}")
    print(f"  COMPARISON")
    print(f"{'='*70}")
    
    # Create lookup by text content
    orig_by_text = {t['text']: t for t in orig_texts}
    new_by_text = {t['text']: t for t in new_texts}
    
    # Find missing texts
    missing = []
    for t in orig_texts:
        if t['text'] not in new_by_text:
            missing.append(t)
    
    # Find new texts
    added = []
    for t in new_texts:
        if t['text'] not in orig_by_text:
            added.append(t)
    
    if missing:
        print(f"\n  MISSING in re-parsed ({len(missing)}):")
        for t in missing:
            print(f"    - [{t['location']}] '{t['text']}' (len={t['full_len']})")
            print(f"      xml_ref: {t['xml_ref']}")
    
    if added:
        print(f"\n  ADDED in re-parsed ({len(added)}):")
        for t in added:
            print(f"    - [{t['location']}] '{t['text']}' (len={t['full_len']})")
            print(f"      xml_ref: {t['xml_ref']}")
    
    if not missing and not added:
        print("\n  ✓ All texts match!")


def main():
    print("Debugging text loss in roundtrip...")
    
    # Parse original
    orig_doc = docx_to_json(str(TEST_FILE), "original")
    orig_texts = extract_all_text_detailed(orig_doc, "ORIGINAL")
    
    # Export and re-parse
    if not EXPORT_FILE.exists():
        apply_json_to_docx(orig_doc, str(TEST_FILE), str(EXPORT_FILE))
    
    reparsed_doc = docx_to_json(str(EXPORT_FILE), "reparsed")
    new_texts = extract_all_text_detailed(reparsed_doc, "RE-PARSED")
    
    # Compare
    compare_texts(orig_texts, new_texts)


if __name__ == "__main__":
    main()
