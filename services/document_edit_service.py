"""Document Edit Service - Thin orchestration layer for AI-powered edits.

This service is the single entry point for AI document editing.
It handles:
1. Locating target block/cell and extracting text
2. Calling the AI agent (Gemini + LangGraph)
3. Applying edited text back to runs
4. Running evals on every edit (baked-in, not bolt-on)
5. Validation

Flow:
    HTTP → Route → DocumentEditService → DocumentEditAgent → Gemini
                                       → EditEvaluator (baked-in)
                                       → Validation
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import List

from models.schemas import BlockType, DocumentJSON, Run
from services.ai_agent import DocumentEditAgent, EditEvaluator, get_edit_agent
from services.ai_config import get_ai_settings
from services.document_engine import validate_document_json

logger = logging.getLogger(__name__)


@dataclass
class EditTarget:
    """Represents the target of an edit operation."""
    found: bool
    original_text: str
    runs: List[Run] | None
    error: str | None = None


@dataclass
class EditResult:
    """Result of an AI edit operation."""
    success: bool
    edited_text: str
    original_text: str
    intent: str
    confidence: float
    reasoning: str
    
    # Eval scores (baked-in)
    eval_scores: dict | None = None
    
    # Validation
    validation_passed: bool = True
    validation_errors: list[str] | None = None
    
    # Error info
    error: str | None = None


class DocumentEditService:
    """Orchestrates AI-powered document edits.
    
    Single responsibility: coordinate the edit flow.
    Does NOT handle HTTP concerns or DB persistence.
    """
    
    def __init__(self, agent: DocumentEditAgent | None = None):
        self._agent = agent or get_edit_agent()
        self._settings = get_ai_settings()
    
    def locate_edit_target(
        self,
        doc: DocumentJSON,
        block_id: str,
        cell_id: str | None = None,
    ) -> EditTarget:
        """Locate the target block or cell for editing.
        
        Args:
            doc: The document to search
            block_id: ID of the block to edit
            cell_id: Optional cell ID for table cell edits
            
        Returns:
            EditTarget with found=True and runs if successful,
            or found=False with error message if not.
        """
        # Check if editing a table cell (handles nested tables)
        if cell_id:
            def find_cell_in_table(table_block) -> EditTarget | None:
                """Recursively search for cell in table and nested tables."""
                for row in table_block.rows:
                    for cell in row.cells:
                        if cell.id == cell_id and cell.blocks:
                            # Find first paragraph block
                            para = next(
                                (b for b in cell.blocks if getattr(b, "type", None) == BlockType.PARAGRAPH),
                                None
                            )
                            if para:
                                original_text = "".join(r.text or "" for r in para.runs)
                                return EditTarget(
                                    found=True,
                                    original_text=original_text,
                                    runs=para.runs,
                                )
                        # Search nested tables
                        for nested_block in cell.blocks:
                            if getattr(nested_block, "type", None) == BlockType.TABLE:
                                result = find_cell_in_table(nested_block)
                                if result and result.found:
                                    return result
                return None
            
            for block in doc.blocks:
                if getattr(block, "type", None) != BlockType.TABLE:
                    continue
                result = find_cell_in_table(block)
                if result and result.found:
                    return result
            
            return EditTarget(
                found=False,
                original_text="",
                runs=None,
                error=f"Table cell '{cell_id}' not found",
            )
        
        # Edit a paragraph block
        for block in doc.blocks:
            if (
                getattr(block, "id", None) == block_id
                and getattr(block, "type", None) == BlockType.PARAGRAPH
            ):
                original_text = "".join(r.text or "" for r in block.runs)
                return EditTarget(
                    found=True,
                    original_text=original_text,
                    runs=block.runs,
                )
        
        return EditTarget(
            found=False,
            original_text="",
            runs=None,
            error=f"Block '{block_id}' not found or not editable",
        )
    
    async def apply_ai_edit(
        self,
        doc: DocumentJSON,
        block_id: str,
        instruction: str,
        cell_id: str | None = None,
        context: str = "",
    ) -> EditResult:
        """Apply an AI edit to a document block or cell.
        
        This is the main entry point for AI edits.
        Evals are run automatically on every edit.
        
        Args:
            doc: The document to edit (will be mutated)
            block_id: ID of the block to edit
            instruction: The edit instruction (e.g., "make more formal")
            cell_id: Optional cell ID for table cell edits
            context: Optional additional context for the AI
            
        Returns:
            EditResult with success status, edited text, and eval scores.
        """
        # Step 1: Locate target
        target = self.locate_edit_target(doc, block_id, cell_id)
        if not target.found:
            return EditResult(
                success=False,
                edited_text="",
                original_text="",
                intent="",
                confidence=0.0,
                reasoning="",
                error=target.error,
            )
        
        # Step 2: Call AI agent
        try:
            ai_result = await self._agent.edit(
                text=target.original_text,
                instruction=instruction,
                context=context,
            )
        except Exception as e:
            logger.error(f"AI agent failed: {e}")
            return EditResult(
                success=False,
                edited_text=target.original_text,
                original_text=target.original_text,
                intent="",
                confidence=0.0,
                reasoning="",
                error=f"AI edit failed: {str(e)}",
            )
        
        edited_text = ai_result["edited_text"]
        intent = ai_result.get("intent", "other")
        confidence = ai_result.get("confidence", 0.0)
        reasoning = ai_result.get("reasoning", "")
        
        # Step 3: Run evals (BAKED-IN, not optional)
        eval_scores = None
        if self._settings.evals_enabled:
            eval_scores = EditEvaluator.evaluate_edit(
                original=target.original_text,
                edited=edited_text,
                instruction=instruction,
                intent=intent,
            )
            
            if self._settings.evals_log_to_console:
                logger.info(
                    f"AI Edit: intent={intent}, confidence={confidence:.2f}, "
                    f"eval_overall={eval_scores['overall_score']:.2f}"
                )
        
        # Step 4: Apply edit to runs
        if target.runs:
            if len(target.runs) > 0:
                target.runs[0].text = edited_text
                # Clear other runs (consolidate into first)
                for run in target.runs[1:]:
                    run.text = ""
        
        # Step 5: Validate document
        validation = validate_document_json(doc)
        if not validation.is_valid:
            # Rollback the edit
            if target.runs and len(target.runs) > 0:
                target.runs[0].text = target.original_text
            
            return EditResult(
                success=False,
                edited_text=edited_text,
                original_text=target.original_text,
                intent=intent,
                confidence=confidence,
                reasoning=reasoning,
                eval_scores=eval_scores,
                validation_passed=False,
                validation_errors=[e.message for e in validation.errors],
                error="Validation failed after AI edit",
            )
        
        return EditResult(
            success=True,
            edited_text=edited_text,
            original_text=target.original_text,
            intent=intent,
            confidence=confidence,
            reasoning=reasoning,
            eval_scores=eval_scores,
            validation_passed=True,
        )


# Singleton instance
_service: DocumentEditService | None = None


def get_document_edit_service() -> DocumentEditService:
    """Get the document edit service singleton."""
    global _service
    if _service is None:
        _service = DocumentEditService()
    return _service
