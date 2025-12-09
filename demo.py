#!/usr/bin/env python3
"""
DOCX Pipeline Interactive Demo

A guided tour through the Document API for developers.
Run with: python demo.py
"""

import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

try:
    import requests
except ImportError:
    print("Missing 'requests' library. Install with: pip install requests")
    sys.exit(1)

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.syntax import Syntax
    from rich.prompt import Prompt, Confirm
    from rich.text import Text
    from rich.box import ROUNDED, DOUBLE
    from rich.markdown import Markdown
except ImportError:
    print("Missing 'rich' library. Install with: pip install rich")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

BASE_URL = "http://127.0.0.1:8000"
UPLOADS_DIR = Path("data/uploads/docx")
console = Console()


# ─────────────────────────────────────────────────────────────────────────────
# Data Classes
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class DemoStep:
    """Represents a single step in the demo workflow."""
    title: str
    description: str
    method: str
    endpoint_template: str
    body_template: Optional[dict] = None
    content_type: str = "application/json"
    is_file_upload: bool = False
    is_file_download: bool = False
    is_manual_edit: bool = False  # For PUT /documents/{id} - full JSON edit
    success_message: str = "✓ Success!"
    next_hint: str = ""
    editable_params: list = field(default_factory=list)


@dataclass
class DemoSession:
    """Tracks the current demo session state."""
    selected_file: Optional[str] = None
    document_id: Optional[str] = None
    last_response: Optional[dict] = None
    blocks: list = field(default_factory=list)
    checkboxes: list = field(default_factory=list)
    dropdowns: list = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# Demo Steps Definition
# ─────────────────────────────────────────────────────────────────────────────

DEMO_STEPS = [
    DemoStep(
        title="Upload Document",
        description="Upload a DOCX file to the server and parse it into JSON structure.",
        method="POST",
        endpoint_template="/documents/",
        is_file_upload=True,
        success_message="✓ Document uploaded and parsed successfully!",
        next_hint="Next: View the parsed JSON structure",
    ),
    DemoStep(
        title="Get Document JSON",
        description="Retrieve the full JSON structure of the uploaded document.",
        method="GET",
        endpoint_template="/documents/{document_id}",
        success_message="✓ Document JSON retrieved!",
        next_hint="Next: Validate the document parsing",
    ),
    DemoStep(
        title="Validate Document",
        description="Validate the document against the original DOCX to check parsing fidelity.",
        method="GET",
        endpoint_template="/documents/{document_id}/validate",
        success_message="✓ Validation complete!",
        next_hint="Next: Try editing the document (manual or AI)",
    ),
    DemoStep(
        title="Update Document (Manual Edit)",
        description="Update the document JSON directly. Modify text, runs, or structure manually.",
        method="PUT",
        endpoint_template="/documents/{document_id}",
        body_template=None,  # Special handling - uses full document JSON
        editable_params=["block_index", "run_index", "new_text"],
        success_message="✓ Document updated!",
        next_hint="Next: Try form fields or AI edit",
        is_manual_edit=True,
    ),
    DemoStep(
        title="Update Checkbox",
        description="Toggle a checkbox form field in the document.",
        method="POST",
        endpoint_template="/documents/{document_id}/checkbox",
        body_template={"checkbox_id": "{checkbox_id}", "checked": True},
        editable_params=["checkbox_id", "checked"],
        success_message="✓ Checkbox updated!",
        next_hint="Next: Try updating a dropdown or export",
    ),
    DemoStep(
        title="Update Dropdown",
        description="Select a value in a dropdown form field.",
        method="POST",
        endpoint_template="/documents/{document_id}/dropdown",
        body_template={"dropdown_id": "{dropdown_id}", "selected": "{selected_value}"},
        editable_params=["dropdown_id", "selected_value"],
        success_message="✓ Dropdown updated!",
        next_hint="Next: Try AI edit or export",
    ),
    DemoStep(
        title="AI Edit Block",
        description="Use AI to edit a paragraph or table cell with natural language instructions.",
        method="POST",
        endpoint_template="/documents/{document_id}/ai-edit",
        body_template={"block_id": "{block_id}", "instruction": "{instruction}"},
        editable_params=["block_id", "instruction"],
        success_message="✓ AI edit applied!",
        next_hint="Next: Validate and export",
    ),
    DemoStep(
        title="Validate Export (Dry Run)",
        description="Test the export process without saving the file.",
        method="POST",
        endpoint_template="/documents/{document_id}/validate-export",
        success_message="✓ Export validation complete!",
        next_hint="Next: Export the document",
    ),
    DemoStep(
        title="Export to DOCX",
        description="Export the JSON back to a DOCX file.",
        method="POST",
        endpoint_template="/documents/{document_id}/export",
        success_message="✓ Document exported!",
        next_hint="Next: Download the exported file",
    ),
    DemoStep(
        title="Download Exported File",
        description="Download the reconstructed DOCX file to your local directory.",
        method="POST",
        endpoint_template="/documents/{document_id}/export/file",
        is_file_download=True,
        success_message="✓ File downloaded!",
        next_hint="Next: View HTML preview",
    ),
    DemoStep(
        title="HTML Preview",
        description="Get an HTML preview of the document (can open in browser).",
        method="GET",
        endpoint_template="/documents/{document_id}/preview/html",
        success_message="✓ HTML preview generated!",
        next_hint="Demo complete! Try other files or explore more endpoints.",
    ),
]


# ─────────────────────────────────────────────────────────────────────────────
# Helper Functions
# ─────────────────────────────────────────────────────────────────────────────

def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def check_server() -> bool:
    """Check if the server is running."""
    try:
        response = requests.get(f"{BASE_URL}/", timeout=3)
        return response.status_code == 200
    except requests.exceptions.ConnectionError:
        return False
    except Exception:
        return False


def get_available_files() -> list[tuple[str, int]]:
    """Get list of DOCX files with their sizes."""
    files = []
    if UPLOADS_DIR.exists():
        for f in UPLOADS_DIR.iterdir():
            if f.suffix.lower() == ".docx":
                files.append((f.name, f.stat().st_size))
    return sorted(files, key=lambda x: x[1])  # Sort by size


def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes // 1024} KB"
    else:
        return f"{size_bytes // (1024 * 1024)} MB"


def format_json(data: Any, max_lines: int = 30) -> str:
    """Format JSON with truncation for large responses."""
    formatted = json.dumps(data, indent=2)
    lines = formatted.split('\n')
    if len(lines) > max_lines:
        half = max_lines // 2
        truncated = lines[:half] + [f"  ... ({len(lines) - max_lines} lines hidden) ..."] + lines[-half:]
        return '\n'.join(truncated)
    return formatted


def build_curl_command(step: DemoStep, session: DemoSession, params: dict) -> str:
    """Build the equivalent curl command for display."""
    endpoint = step.endpoint_template.format(document_id=session.document_id or "{document_id}")
    url = f"{BASE_URL}{endpoint}"
    
    if step.is_file_upload:
        return f'curl.exe -X POST "{url}" \\\n  -F "file=@data/uploads/docx/{session.selected_file}"'
    
    if step.is_file_download:
        output_name = f"{Path(session.selected_file).stem}_exported.docx" if session.selected_file else "exported.docx"
        return f'curl.exe -X POST "{url}" \\\n  --output {output_name}'
    
    if step.is_manual_edit:
        # Show the workflow for manual edit
        return (
            f'# Step 1: Get current JSON\n'
            f'curl.exe -X GET "{BASE_URL}/documents/{session.document_id or "{document_id}"}" -o document.json\n\n'
            f'# Step 2: Edit document.json in your editor\n'
            f'# (modify blocks[{params.get("block_index", 0)}].runs[{params.get("run_index", 0)}].text)\n\n'
            f'# Step 3: Upload modified JSON\n'
            f'curl.exe -X PUT "{url}" \\\n  -H "Content-Type: application/json" \\\n  -d @document.json'
        )
    
    if step.body_template:
        body = json.dumps(params, indent=2)
        return f'curl.exe -X {step.method} "{url}" \\\n  -H "Content-Type: application/json" \\\n  -d \'{body}\''
    
    return f'curl.exe -X {step.method} "{url}"'


def execute_step(step: DemoStep, session: DemoSession, params: dict) -> tuple[bool, Any, int]:
    """Execute an API call and return (success, response_data, status_code)."""
    endpoint = step.endpoint_template.format(document_id=session.document_id)
    url = f"{BASE_URL}{endpoint}"
    
    try:
        # Manual edit: modify a specific run in the document
        if step.is_manual_edit:
            # First get current document
            get_response = requests.get(f"{BASE_URL}/documents/{session.document_id}", timeout=30)
            if get_response.status_code != 200:
                return False, {"error": "Failed to fetch document for editing"}, get_response.status_code
            
            doc = get_response.json()
            
            # Apply the edit
            block_idx = int(params.get("block_index", 0))
            run_idx = int(params.get("run_index", 0))
            new_text = params.get("new_text", "")
            
            if block_idx < len(doc.get("blocks", [])):
                block = doc["blocks"][block_idx]
                if block.get("type") == "paragraph" and run_idx < len(block.get("runs", [])):
                    old_text = block["runs"][run_idx].get("text", "")
                    block["runs"][run_idx]["text"] = new_text
                    
                    # PUT the updated document
                    response = requests.put(url, json=doc, timeout=30)
                    
                    result_data = {
                        "edit_applied": True,
                        "block_index": block_idx,
                        "run_index": run_idx,
                        "old_text": old_text,
                        "new_text": new_text,
                    }
                    if response.status_code < 400:
                        return True, result_data, response.status_code
                    else:
                        try:
                            result_data["error"] = response.json()
                        except:
                            result_data["error"] = response.text
                        return False, result_data, response.status_code
                else:
                    return False, {"error": f"Run index {run_idx} out of range or block is not a paragraph"}, 400
            else:
                return False, {"error": f"Block index {block_idx} out of range"}, 400
        
        if step.is_file_upload:
            file_path = UPLOADS_DIR / session.selected_file
            with open(file_path, "rb") as f:
                response = requests.post(url, files={"file": f}, timeout=60)
        elif step.is_file_download:
            response = requests.post(url, timeout=60)
            if response.status_code == 200:
                output_name = f"{Path(session.selected_file).stem}_exported.docx"
                with open(output_name, "wb") as f:
                    f.write(response.content)
                return True, {"message": f"File saved as {output_name}", "size": len(response.content)}, 200
        elif step.method == "GET":
            response = requests.get(url, timeout=30)
        elif step.method == "POST":
            if step.body_template:
                response = requests.post(url, json=params, timeout=60)
            else:
                response = requests.post(url, timeout=60)
        elif step.method == "PUT":
            response = requests.put(url, json=params, timeout=60)
        else:
            return False, {"error": f"Unknown method: {step.method}"}, 0
        
        # Try to parse JSON response
        try:
            data = response.json()
        except json.JSONDecodeError:
            # For HTML responses or non-JSON
            if "text/html" in response.headers.get("content-type", ""):
                data = {"html_preview": f"HTML content ({len(response.text)} chars)", 
                        "url": f"{url} (open in browser)"}
            else:
                data = {"raw": response.text[:500]}
        
        return response.status_code < 400, data, response.status_code
        
    except requests.exceptions.ConnectionError:
        return False, {"error": "Connection failed. Is the server running?"}, 0
    except requests.exceptions.Timeout:
        return False, {"error": "Request timed out"}, 0
    except Exception as e:
        return False, {"error": str(e)}, 0


# ─────────────────────────────────────────────────────────────────────────────
# UI Components
# ─────────────────────────────────────────────────────────────────────────────

def show_header():
    """Display the demo header."""
    console.print()
    console.print(Panel(
        "[bold cyan]DOCX Pipeline Interactive Demo[/bold cyan]\n"
        "[dim]Learn how to use the Document API step by step[/dim]",
        box=DOUBLE,
        border_style="cyan",
        padding=(1, 2),
    ))


def show_server_status():
    """Check and display server status."""
    console.print("\n[dim]Checking server status...[/dim]")
    if check_server():
        console.print(f"[green]✓ Server is running at {BASE_URL}[/green]\n")
        return True
    else:
        console.print(f"[red]✗ Server is not running at {BASE_URL}[/red]")
        console.print("\n[yellow]Start the server with:[/yellow]")
        console.print(Panel(
            ".venv\\Scripts\\Activate.ps1\nuvicorn main:app --reload --port 8000",
            title="Command",
            border_style="yellow",
        ))
        return False


def show_file_selection() -> Optional[str]:
    """Display file selection menu and return selected filename."""
    files = get_available_files()
    
    if not files:
        console.print("[red]No DOCX files found in data/uploads/docx/[/red]")
        return None
    
    table = Table(title="Available DOCX Files", box=ROUNDED, border_style="blue")
    table.add_column("#", style="cyan", justify="right")
    table.add_column("Filename", style="white")
    table.add_column("Size", style="dim", justify="right")
    table.add_column("", style="yellow")
    
    # Find smallest file for recommendation
    smallest_idx = 0
    smallest_size = files[0][1]
    for i, (name, size) in enumerate(files):
        if size < smallest_size:
            smallest_size = size
            smallest_idx = i
    
    for i, (name, size) in enumerate(files):
        marker = "⭐ Recommended" if i == smallest_idx else ""
        table.add_row(str(i + 1), name, format_size(size), marker)
    
    console.print(table)
    console.print()
    
    while True:
        choice = Prompt.ask(
            f"Select file [cyan][1-{len(files)}][/cyan] or [red][q][/red] to quit",
            default="1"
        )
        
        if choice.lower() == 'q':
            return None
        
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx][0]
            console.print("[red]Invalid selection[/red]")
        except ValueError:
            console.print("[red]Please enter a number[/red]")


def show_step_menu() -> str:
    """Show the step selection menu."""
    console.print("\n[bold]Available Steps:[/bold]")
    
    table = Table(box=ROUNDED, border_style="dim")
    table.add_column("#", style="cyan", justify="right", width=3)
    table.add_column("Step", style="white")
    table.add_column("Description", style="dim")
    
    for i, step in enumerate(DEMO_STEPS):
        table.add_row(str(i + 1), step.title, step.description[:50] + "...")
    
    console.print(table)
    console.print()
    console.print("[dim]Quick paths: [cyan][a][/cyan] All steps | [cyan][q][/cyan] Quick (upload→export→download) | [cyan][f][/cyan] Select new file[/dim]")
    
    return Prompt.ask(
        f"Select step [cyan][1-{len(DEMO_STEPS)}][/cyan], path, or [red][x][/red] to exit",
        default="1"
    )


def show_step_detail(step: DemoStep, step_num: int, total: int, session: DemoSession, params: dict):
    """Display a single step with its command."""
    console.print()
    console.print(Panel(
        f"[bold]{step.title}[/bold]\n\n{step.description}",
        title=f"Step {step_num}/{total}",
        border_style="green",
    ))
    
    # Show the curl command
    curl_cmd = build_curl_command(step, session, params)
    console.print("\n[bold]Command:[/bold]")
    console.print(Syntax(curl_cmd, "bash", theme="monokai", line_numbers=False, word_wrap=True))


def edit_params(step: DemoStep, session: DemoSession, current_params: dict) -> dict:
    """Allow user to edit step parameters."""
    if not step.editable_params:
        console.print("[dim]No editable parameters for this step[/dim]")
        return current_params
    
    console.print("\n[bold]Edit Parameters:[/bold]")
    new_params = current_params.copy()
    
    for param in step.editable_params:
        current_val = current_params.get(param, "")
        
        # Provide suggestions based on session context
        suggestions = ""
        if param == "checkbox_id" and session.checkboxes:
            suggestions = f" [dim](available: {', '.join(c.get('id', 'unknown')[:20] for c in session.checkboxes[:3])})[/dim]"
        elif param == "dropdown_id" and session.dropdowns:
            suggestions = f" [dim](available: {', '.join(d.get('id', 'unknown')[:20] for d in session.dropdowns[:3])})[/dim]"
        elif param == "block_id" and session.blocks:
            block_ids = [b.get('id', '') for b in session.blocks[:5] if b.get('id')]
            suggestions = f" [dim](available: {', '.join(block_ids)})[/dim]"
        elif param == "checked":
            suggestions = " [dim](true/false)[/dim]"
        elif param == "block_index":
            para_indices = [str(i) for i, b in enumerate(session.blocks[:10]) if b.get('type') == 'paragraph']
            suggestions = f" [dim](paragraph blocks: {', '.join(para_indices)})[/dim]"
        elif param == "run_index":
            suggestions = " [dim](usually 0 for first text run)[/dim]"
        elif param == "new_text":
            suggestions = " [dim](the new text content)[/dim]"
        
        new_val = Prompt.ask(f"  {param}{suggestions}", default=str(current_val))
        
        # Type conversion
        if param == "checked":
            new_params[param] = new_val.lower() in ("true", "1", "yes")
        elif param in ("block_index", "run_index"):
            try:
                new_params[param] = int(new_val)
            except ValueError:
                new_params[param] = 0
        else:
            new_params[param] = new_val
    
    return new_params


def show_response(success: bool, data: Any, status_code: int, step: DemoStep):
    """Display the API response."""
    if success:
        console.print(f"\n[green]Response ({status_code} OK):[/green]")
    else:
        console.print(f"\n[red]Response ({status_code} Error):[/red]")
    
    formatted = format_json(data)
    console.print(Syntax(formatted, "json", theme="monokai", line_numbers=False, word_wrap=True))
    
    if success:
        console.print(f"\n[green]{step.success_message}[/green]")
        if step.next_hint:
            console.print(f"[dim]{step.next_hint}[/dim]")
    else:
        console.print("\n[yellow]Troubleshooting:[/yellow]")
        if status_code == 404:
            console.print("  • Document not found - did you upload it first?")
        elif status_code == 0:
            console.print("  • Server connection failed - is the server running?")
        else:
            console.print("  • Check the error message above for details")


def extract_session_data(data: dict, session: DemoSession):
    """Extract useful data from response into session."""
    if "id" in data:
        session.document_id = data["id"]
    if "blocks" in data:
        session.blocks = data["blocks"]
    if "checkboxes" in data:
        session.checkboxes = data["checkboxes"]
    if "dropdowns" in data:
        session.dropdowns = data["dropdowns"]
    session.last_response = data


def get_default_params(step: DemoStep, session: DemoSession) -> dict:
    """Get default parameters for a step based on session context."""
    params = {}
    
    # Special handling for manual edit
    if step.is_manual_edit:
        # Find first paragraph block with text
        default_text = "[EDITED BY DEMO]"
        for i, block in enumerate(session.blocks):
            if block.get("type") == "paragraph":
                runs = block.get("runs", [])
                if runs and runs[0].get("text"):
                    original_text = runs[0].get("text", "")
                    return {
                        "block_index": i,
                        "run_index": 0,
                        "new_text": f"{original_text} [EDITED]",
                    }
        return {"block_index": 0, "run_index": 0, "new_text": default_text}
    
    if step.body_template:
        for key, val in step.body_template.items():
            if isinstance(val, str) and val.startswith("{") and val.endswith("}"):
                param_name = val[1:-1]
                # Provide smart defaults
                if param_name == "checkbox_id" and session.checkboxes:
                    params[key] = session.checkboxes[0].get("id", "checkbox-0")
                elif param_name == "dropdown_id" and session.dropdowns:
                    params[key] = session.dropdowns[0].get("id", "dropdown-0")
                elif param_name == "selected_value" and session.dropdowns:
                    options = session.dropdowns[0].get("options", ["Option1"])
                    params[key] = options[0] if options else "Yes"
                elif param_name == "block_id" and session.blocks:
                    params[key] = session.blocks[0].get("id", "p-0")
                elif param_name == "instruction":
                    params[key] = "make this more formal"
                else:
                    params[key] = ""
            else:
                params[key] = val
    
    return params


 # ─────────────────────────────────────────────────────────────────────────────
 # Main Demo Flow (Linear Wizard)
 # ─────────────────────────────────────────────────────────────────────────────


def run_linear_workflow(session: DemoSession):
    """Run the full demo as a linear wizard after file selection.

    Behaviour:
    - Walks through DEMO_STEPS sequentially.
    - Each step shows the command, allows constrained edits, and can be
      repeated before moving on.
    - Navigation is local to the flow: [r]un, [e]dit, [n]ext, [b]ack, [q]uit.
    """

    step_index = 0
    total_steps = len(DEMO_STEPS)

    while 0 <= step_index < total_steps:
        step = DEMO_STEPS[step_index]

        # Skip form-field steps if the current document has no such fields
        if step.title == "Update Checkbox" and not session.checkboxes:
            step_index += 1
            continue
        if step.title == "Update Dropdown" and not session.dropdowns:
            step_index += 1
            continue

        params = get_default_params(step, session)

        while True:
            clear_screen()
            show_header()
            console.print(
                f"\n[dim]File: [cyan]{session.selected_file}[/cyan] | "
                f"Document ID: [cyan]{session.document_id or 'Not uploaded yet'}[/cyan] | "
                f"Step {step_index + 1}/{total_steps}[/dim]"
            )

            show_step_detail(step, step_index + 1, total_steps, session, params)

            console.print("\n[bold]Actions:[/bold]")
            console.print("  [cyan]1.[/cyan] Run command")
            if step.editable_params:
                console.print("  [cyan]2.[/cyan] Edit allowed parameters")
                console.print("  [cyan]3.[/cyan] Next step")
                next_index = 3
            else:
                console.print("  [cyan]2.[/cyan] Next step")
                next_index = 2

            back_index = next_index + 1 if step_index > 0 else None
            quit_index = (back_index or next_index) + 1

            if step_index > 0:
                console.print(f"  [cyan]{back_index}.[/cyan] Back to previous step")
            console.print(f"  [red]{quit_index}.[/red] Quit demo")

            action_raw = Prompt.ask("\nYour choice", default="1").lower()

            # Map numeric choices to actions, still accept letter shortcuts
            if action_raw in ("1", "r"):
                action = "r"
            elif step.editable_params and action_raw in ("2", "e"):
                action = "e"
            elif (not step.editable_params and action_raw in ("2", "n")) or (
                step.editable_params and action_raw in ("3", "n")
            ):
                action = "n"
            elif back_index is not None and action_raw in (str(back_index), "b"):
                action = "b"
            elif action_raw in (str(quit_index), "q"):
                action = "q"
            else:
                # Unknown input, re-prompt this step
                continue

            if action == "q":
                return
            if action == "b" and step_index > 0:
                step_index -= 1
                break
            if action == "n":
                step_index += 1
                break
            if action == "e" and step.editable_params:
                params = edit_params(step, session, params)
                continue
            if action == "r":
                console.print("\n[dim]Executing...[/dim] ⏳")
                success, data, status_code = execute_step(step, session, params)

                if success:
                    extract_session_data(data, session)

                show_response(success, data, status_code, step)

                console.print()
                follow_raw = Prompt.ask(
                    "1. Next step  |  2. Edit & rerun  |  3. Back  |  4. Quit",
                    default="1",
                ).lower()

                if follow_raw in ("1", "n"):
                    follow_up = "n"
                elif follow_raw in ("2", "e"):
                    follow_up = "e"
                elif follow_raw in ("3", "b"):
                    follow_up = "b"
                elif follow_raw in ("4", "q"):
                    follow_up = "q"
                else:
                    # On invalid choice, default to next step
                    follow_up = "n"

                if follow_up == "q":
                    return
                if follow_up == "b" and step_index > 0:
                    step_index -= 1
                    break
                if follow_up == "e" and step.editable_params:
                    params = edit_params(step, session, params)
                    continue

                # default: move to next step
                step_index += 1
                break


def main():
    """Main entry point: choose file, then run linear wizard."""
    clear_screen()
    show_header()

    # Check server
    if not show_server_status():
        if not Confirm.ask("\nContinue anyway?", default=False):
            console.print("[dim]Goodbye![/dim]")
            return

    # File selection
    selected_file = show_file_selection()
    if not selected_file:
        console.print("[dim]Goodbye![/dim]")
        return

    session = DemoSession(selected_file=selected_file)

    # Run the linear guided tour
    run_linear_workflow(session)

    console.print("\n[dim]Demo finished. You can run it again with another file any time.[/dim]")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n[dim]Interrupted. Goodbye![/dim]")
