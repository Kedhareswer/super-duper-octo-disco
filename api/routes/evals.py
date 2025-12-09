"""API routes for AI evaluation dashboard."""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional
from dataclasses import dataclass, field, asdict
import json

from fastapi import APIRouter
from pydantic import BaseModel

from services.ai_agent import EditEvaluator, get_edit_agent


router = APIRouter(prefix="/evals", tags=["evals"])


# In-memory storage for eval results (would use DB in production)
_eval_history: List[dict] = []


class EvalRequest(BaseModel):
    """Request to evaluate an edit."""
    original_text: str
    edited_text: str
    instruction: str
    intent: str


class EvalResult(BaseModel):
    """Result of an evaluation."""
    preservation_score: float
    instruction_adherence: float
    fluency_score: float
    overall_score: float


class TestCaseRequest(BaseModel):
    """Request to run a test case."""
    original: str
    instruction: str


class TestCaseResult(BaseModel):
    """Result of a test case."""
    original: str
    instruction: str
    edited: str
    intent: str
    confidence: float
    eval_scores: EvalResult
    timestamp: str


class DashboardStats(BaseModel):
    """Statistics for the eval dashboard."""
    total_evals: int
    avg_overall_score: float
    avg_preservation: float
    avg_adherence: float
    avg_fluency: float
    recent_evals: List[dict]


@router.post("/evaluate", response_model=EvalResult)
async def evaluate_edit(request: EvalRequest) -> EvalResult:
    """Evaluate the quality of an edit."""
    metrics = EditEvaluator.evaluate_edit(
        request.original_text,
        request.edited_text,
        request.instruction,
        request.intent,
    )
    
    result = EvalResult(**metrics)
    
    # Store in history
    _eval_history.append({
        "original": request.original_text[:100],
        "edited": request.edited_text[:100],
        "instruction": request.instruction,
        "intent": request.intent,
        "scores": metrics,
        "timestamp": datetime.now().isoformat(),
    })
    
    # Keep only last 100 evals
    if len(_eval_history) > 100:
        _eval_history.pop(0)
    
    return result


@router.post("/test", response_model=TestCaseResult)
async def run_test_case(request: TestCaseRequest) -> TestCaseResult:
    """Run a single test case through the AI agent and evaluate it."""
    agent = get_edit_agent()
    
    # Run the edit
    result = await agent.edit(request.original, request.instruction)
    
    # Evaluate the result
    metrics = EditEvaluator.evaluate_edit(
        request.original,
        result["edited_text"],
        request.instruction,
        result["intent"],
    )
    
    test_result = TestCaseResult(
        original=request.original,
        instruction=request.instruction,
        edited=result["edited_text"],
        intent=result["intent"],
        confidence=result["confidence"],
        eval_scores=EvalResult(**metrics),
        timestamp=datetime.now().isoformat(),
    )
    
    # Store in history
    _eval_history.append({
        "original": request.original[:100],
        "edited": result["edited_text"][:100],
        "instruction": request.instruction,
        "intent": result["intent"],
        "confidence": result["confidence"],
        "scores": metrics,
        "timestamp": datetime.now().isoformat(),
    })
    
    return test_result


@router.post("/test-suite")
async def run_test_suite():
    """Run the built-in test suite."""
    results = EditEvaluator.run_test_suite()
    
    passed = sum(1 for r in results if r["passed"])
    total = len(results)
    
    return {
        "passed": passed,
        "total": total,
        "pass_rate": passed / total if total > 0 else 0,
        "results": results,
    }


@router.get("/dashboard", response_model=DashboardStats)
async def get_dashboard_stats() -> DashboardStats:
    """Get statistics for the eval dashboard."""
    if not _eval_history:
        return DashboardStats(
            total_evals=0,
            avg_overall_score=0,
            avg_preservation=0,
            avg_adherence=0,
            avg_fluency=0,
            recent_evals=[],
        )
    
    # Calculate averages
    scores = [e["scores"] for e in _eval_history if "scores" in e]
    
    avg_overall = sum(s["overall_score"] for s in scores) / len(scores) if scores else 0
    avg_preservation = sum(s["preservation_score"] for s in scores) / len(scores) if scores else 0
    avg_adherence = sum(s["instruction_adherence"] for s in scores) / len(scores) if scores else 0
    avg_fluency = sum(s["fluency_score"] for s in scores) / len(scores) if scores else 0
    
    return DashboardStats(
        total_evals=len(_eval_history),
        avg_overall_score=round(avg_overall, 3),
        avg_preservation=round(avg_preservation, 3),
        avg_adherence=round(avg_adherence, 3),
        avg_fluency=round(avg_fluency, 3),
        recent_evals=_eval_history[-10:][::-1],  # Last 10, newest first
    )


@router.delete("/history")
async def clear_history():
    """Clear the eval history."""
    global _eval_history
    _eval_history = []
    return {"status": "cleared"}
