"""Debug test6.xlsx parser issue with detailed tracing."""

import sys
import zipfile
import traceback
from pathlib import Path
from xml.etree import ElementTree as ET
from io import BytesIO

ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

TEST_FILE = ROOT / "data" / "uploads" / "excel" / "test6.xlsx"

# Import the parser module to trace the issue
from services.excel_engine import parser

def trace_parse():
    """Trace the parsing with detailed error info."""
    print(f"Debugging parser for: {TEST_FILE}")
    
    try:
        result = parser.xlsx_to_json(str(TEST_FILE), "test6")
        print(f"SUCCESS! Parsed {len(result.sheets)} sheets")
    except ET.ParseError as e:
        print(f"\nParseError: {e}")
        print("\nFull traceback:")
        traceback.print_exc()
    except Exception as e:
        print(f"\n{type(e).__name__}: {e}")
        print("\nFull traceback:")
        traceback.print_exc()


if __name__ == "__main__":
    trace_parse()
