"""
LangGraph-based AI Agent for Document Editing.

Architecture:
- State machine with validation nodes
- Gemini as the LLM backend
- Guardrails for input/output validation
- Structured evals for quality assurance
"""
from __future__ import annotations

import os
import re
from typing import Annotated, Literal, TypedDict
from dataclasses import dataclass

# LangGraph imports
try:
    from langgraph.graph import StateGraph, END
    from langgraph.graph.message import add_messages
    HAS_LANGGRAPH = True
except ImportError:
    HAS_LANGGRAPH = False

# Gemini imports
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False


# ============================================================================
# CONFIGURATION
# ============================================================================

# Import centralized config
from services.ai_config import get_ai_settings, AISettings

# Default blocked patterns for guardrails
DEFAULT_BLOCKED_PATTERNS = [
    r'\b(password|secret|api.?key|token)\s*[:=]',  # Credential patterns
    r'\b\d{3}-\d{2}-\d{4}\b',  # SSN pattern
    r'\b\d{16}\b',  # Credit card pattern
]


@dataclass
class AIConfig:
    """Configuration for the AI agent.
    
    Now pulls defaults from centralized ai_config.py.
    """
    model_name: str | None = None  # Will use centralized config if None
    max_output_tokens: int | None = None
    temperature: float | None = None
    # Guardrails
    max_input_length: int | None = None
    max_output_length: int | None = None
    blocked_patterns: list = None
    
    def __post_init__(self):
        # Pull from centralized settings for any None values
        settings = get_ai_settings()
        
        if self.model_name is None:
            self.model_name = settings.gemini.model_name
        if self.max_output_tokens is None:
            self.max_output_tokens = settings.max_output_tokens
        if self.temperature is None:
            self.temperature = settings.temperature
        if self.max_input_length is None:
            self.max_input_length = settings.max_input_length
        if self.max_output_length is None:
            self.max_output_length = settings.max_output_length
        if self.blocked_patterns is None:
            self.blocked_patterns = DEFAULT_BLOCKED_PATTERNS.copy()


# ============================================================================
# STATE DEFINITION
# ============================================================================

class EditState(TypedDict):
    """State for the document editing workflow."""
    # Input
    original_text: str
    instruction: str
    context: str  # Additional context about the document
    
    # Processing
    intent: str  # Parsed intent (e.g., "formalize", "summarize", "correct")
    validation_passed: bool
    validation_errors: list[str]
    
    # Output
    edited_text: str
    confidence: float
    reasoning: str


# ============================================================================
# GUARDRAILS
# ============================================================================

class Guardrails:
    """Input and output validation guardrails."""
    
    def __init__(self, config: AIConfig):
        self.config = config
    
    def validate_input(self, text: str, instruction: str) -> tuple[bool, list[str]]:
        """Validate input before processing."""
        errors = []
        
        # Length check
        if len(text) > self.config.max_input_length:
            errors.append(f"Input text too long ({len(text)} > {self.config.max_input_length})")
        
        if len(instruction) > 500:
            errors.append("Instruction too long (max 500 characters)")
        
        # Empty check
        if not text.strip():
            errors.append("Input text is empty")
        
        if not instruction.strip():
            errors.append("Instruction is empty")
        
        # Blocked pattern check
        combined = f"{text} {instruction}"
        for pattern in self.config.blocked_patterns:
            if re.search(pattern, combined, re.IGNORECASE):
                errors.append(f"Content contains blocked pattern")
                break
        
        return len(errors) == 0, errors
    
    def validate_output(self, original: str, edited: str, instruction: str) -> tuple[bool, list[str]]:
        """Validate output after processing."""
        errors = []
        
        # Length check
        if len(edited) > self.config.max_output_length:
            errors.append(f"Output too long ({len(edited)} > {self.config.max_output_length})")
        
        # Empty output check
        if not edited.strip():
            errors.append("Output is empty")
        
        # Sanity check - output shouldn't be completely different unless summarizing
        if "summar" not in instruction.lower():
            # If output is less than 20% of original, something might be wrong
            if len(edited) < len(original) * 0.2 and len(original) > 50:
                errors.append("Output suspiciously shorter than input")
        
        # Check for hallucinated content markers
        hallucination_markers = [
            "as an ai", "i cannot", "i'm sorry", "i don't have access",
            "based on my training", "as a language model"
        ]
        for marker in hallucination_markers:
            if marker in edited.lower() and marker not in original.lower():
                errors.append("Output contains AI self-reference (likely hallucination)")
                break
        
        return len(errors) == 0, errors


# ============================================================================
# GEMINI CLIENT
# ============================================================================

class GeminiClient:
    """Client for Google's Gemini API.
    
    Uses centralized config from ai_config.py.
    """
    
    def __init__(self, config: AIConfig | None = None):
        self.config = config or AIConfig()
        self._model = None
        self._settings = get_ai_settings()
    
    def _get_model(self):
        """Lazy initialization of Gemini model."""
        if self._model is None:
            # Use centralized config for API key
            api_key = self._settings.gemini.api_key
            if not api_key:
                raise ValueError("GOOGLE_API_KEY or GEMINI_API_KEY environment variable required")
            
            genai.configure(api_key=api_key)
            self._model = genai.GenerativeModel(
                model_name=self.config.model_name,
                generation_config={
                    "temperature": self.config.temperature,
                    "max_output_tokens": self.config.max_output_tokens,
                }
            )
        return self._model
    
    def analyze_intent(self, instruction: str) -> str:
        """Analyze the user's editing intent."""
        model = self._get_model()
        
        prompt = f"""Analyze this document editing instruction and categorize the intent.

INSTRUCTION: {instruction}

Respond with ONE of these categories:
- formalize: Make text more formal/professional
- simplify: Make text simpler/more concise
- correct: Fix grammar, spelling, punctuation
- expand: Add more detail or explanation
- summarize: Condense the text
- rephrase: Reword without changing meaning
- tone_adjust: Change emotional tone
- other: Any other type of edit

Category:"""
        
        response = model.generate_content(prompt)
        intent = response.text.strip().lower()
        
        # Normalize
        valid_intents = ["formalize", "simplify", "correct", "expand", "summarize", "rephrase", "tone_adjust", "other"]
        for valid in valid_intents:
            if valid in intent:
                return valid
        return "other"
    
    def execute_edit(self, text: str, instruction: str, intent: str, context: str = "") -> tuple[str, float, str]:
        """Execute the text edit."""
        model = self._get_model()
        
        context_section = f"\nDOCUMENT CONTEXT: {context}" if context else ""
        
        prompt = f"""You are a document editor assistant. Edit the following text according to the instruction.

INSTRUCTION: {instruction}
EDIT TYPE: {intent}{context_section}

ORIGINAL TEXT:
{text}

RULES:
1. Return ONLY the edited text, nothing else
2. Preserve the original meaning unless explicitly asked to change it
3. Keep formatting markers if present
4. Do not add explanations or commentary
5. If you cannot make the edit, return the original text unchanged

EDITED TEXT:"""
        
        response = model.generate_content(prompt)
        edited = response.text.strip()
        
        # Simple confidence heuristic based on edit magnitude
        original_words = set(text.lower().split())
        edited_words = set(edited.lower().split())
        overlap = len(original_words & edited_words) / max(len(original_words), 1)
        
        # Higher overlap for "correct" type, lower for "summarize"
        if intent == "correct":
            confidence = 0.7 + (overlap * 0.3)
        elif intent == "summarize":
            confidence = 0.8 if len(edited) < len(text) else 0.5
        else:
            confidence = 0.6 + (overlap * 0.3)
        
        reasoning = f"Applied {intent} transformation. Word overlap: {overlap:.2f}"
        
        return edited, min(confidence, 1.0), reasoning


# ============================================================================
# LANGGRAPH WORKFLOW
# ============================================================================

def create_edit_agent(config: AIConfig = None):
    """Create the LangGraph-based editing agent."""
    
    if not HAS_LANGGRAPH:
        raise ImportError("langgraph is required. Install with: pip install langgraph")
    
    if not HAS_GEMINI:
        raise ImportError("google-generativeai is required. Install with: pip install google-generativeai")
    
    config = config or AIConfig()
    guardrails = Guardrails(config)
    gemini = GeminiClient(config)
    
    # Define nodes
    def validate_input_node(state: EditState) -> EditState:
        """Validate input before processing."""
        passed, errors = guardrails.validate_input(
            state["original_text"], 
            state["instruction"]
        )
        return {
            **state,
            "validation_passed": passed,
            "validation_errors": errors,
        }
    
    def analyze_intent_node(state: EditState) -> EditState:
        """Analyze the editing intent."""
        intent = gemini.analyze_intent(state["instruction"])
        return {
            **state,
            "intent": intent,
        }
    
    def execute_edit_node(state: EditState) -> EditState:
        """Execute the edit."""
        edited, confidence, reasoning = gemini.execute_edit(
            state["original_text"],
            state["instruction"],
            state["intent"],
            state.get("context", ""),
        )
        return {
            **state,
            "edited_text": edited,
            "confidence": confidence,
            "reasoning": reasoning,
        }
    
    def validate_output_node(state: EditState) -> EditState:
        """Validate output after processing."""
        passed, errors = guardrails.validate_output(
            state["original_text"],
            state["edited_text"],
            state["instruction"],
        )
        if not passed:
            # Fallback to original if output validation fails
            return {
                **state,
                "edited_text": state["original_text"],
                "validation_passed": False,
                "validation_errors": errors,
                "reasoning": f"Output validation failed: {errors}. Returning original.",
            }
        return {
            **state,
            "validation_passed": True,
            "validation_errors": [],
        }
    
    # Define routing
    def should_continue(state: EditState) -> Literal["analyze", "end"]:
        """Route based on input validation."""
        if state.get("validation_passed", True):
            return "analyze"
        return "end"
    
    # Build graph
    workflow = StateGraph(EditState)
    
    # Add nodes
    workflow.add_node("validate_input", validate_input_node)
    workflow.add_node("analyze_intent", analyze_intent_node)
    workflow.add_node("execute_edit", execute_edit_node)
    workflow.add_node("validate_output", validate_output_node)
    
    # Add edges
    workflow.set_entry_point("validate_input")
    workflow.add_conditional_edges(
        "validate_input",
        should_continue,
        {
            "analyze": "analyze_intent",
            "end": END,
        }
    )
    workflow.add_edge("analyze_intent", "execute_edit")
    workflow.add_edge("execute_edit", "validate_output")
    workflow.add_edge("validate_output", END)
    
    return workflow.compile()


# ============================================================================
# HIGH-LEVEL API
# ============================================================================

class DocumentEditAgent:
    """High-level API for document editing with AI."""
    
    def __init__(self, config: AIConfig = None):
        self.config = config or AIConfig()
        self._agent = None
        self._fallback_mode = False
    
    def _get_agent(self):
        """Lazy initialization of the agent."""
        if self._agent is None:
            try:
                self._agent = create_edit_agent(self.config)
            except (ImportError, ValueError) as e:
                print(f"Warning: Could not initialize LangGraph agent: {e}")
                self._fallback_mode = True
        return self._agent
    
    async def edit(self, text: str, instruction: str, context: str = "") -> dict:
        """
        Edit text according to instruction.
        
        Returns:
            {
                "edited_text": str,
                "confidence": float,
                "reasoning": str,
                "intent": str,
                "validation_errors": list[str],
            }
        """
        # Try LangGraph agent first
        if not self._fallback_mode:
            try:
                agent = self._get_agent()
                if agent:
                    initial_state: EditState = {
                        "original_text": text,
                        "instruction": instruction,
                        "context": context,
                        "intent": "",
                        "validation_passed": True,
                        "validation_errors": [],
                        "edited_text": "",
                        "confidence": 0.0,
                        "reasoning": "",
                    }
                    
                    # Run the graph
                    result = agent.invoke(initial_state)
                    
                    return {
                        "edited_text": result["edited_text"],
                        "confidence": result["confidence"],
                        "reasoning": result["reasoning"],
                        "intent": result["intent"],
                        "validation_errors": result["validation_errors"],
                    }
            except Exception as e:
                print(f"LangGraph agent failed: {e}, falling back to simple mode")
                self._fallback_mode = True
        
        # Fallback to simple Gemini call
        return await self._simple_edit(text, instruction)
    
    async def _simple_edit(self, text: str, instruction: str) -> dict:
        """Simple fallback edit without LangGraph."""
        try:
            gemini = GeminiClient(self.config)
            edited, confidence, reasoning = gemini.execute_edit(text, instruction, "other")
            return {
                "edited_text": edited,
                "confidence": confidence,
                "reasoning": reasoning,
                "intent": "other",
                "validation_errors": [],
            }
        except Exception as e:
            # Ultimate fallback - rule-based
            return self._rule_based_edit(text, instruction)
    
    def _rule_based_edit(self, text: str, instruction: str) -> dict:
        """Rule-based fallback when AI is unavailable."""
        instruction_lower = instruction.lower()
        edited = text
        intent = "other"
        
        if "uppercase" in instruction_lower:
            edited = text.upper()
            intent = "tone_adjust"
        elif "lowercase" in instruction_lower:
            edited = text.lower()
            intent = "tone_adjust"
        elif "formal" in instruction_lower:
            # Simple formalization
            replacements = {
                "don't": "do not", "won't": "will not", "can't": "cannot",
                "shouldn't": "should not", "wouldn't": "would not",
                "isn't": "is not", "aren't": "are not",
            }
            for old, new in replacements.items():
                edited = edited.replace(old, new)
                edited = edited.replace(old.capitalize(), new.capitalize())
            intent = "formalize"
        elif "concise" in instruction_lower or "shorter" in instruction_lower:
            fillers = [" very ", " really ", " quite ", " just ", " actually "]
            for filler in fillers:
                edited = edited.replace(filler, " ")
            intent = "simplify"
        else:
            edited = f"{text} [AI edit requested: {instruction}]"
        
        return {
            "edited_text": edited,
            "confidence": 0.5,
            "reasoning": "Rule-based fallback (no AI available)",
            "intent": intent,
            "validation_errors": [],
        }


# ============================================================================
# EVALS
# ============================================================================

class EditEvaluator:
    """Evaluation framework for edit quality."""
    
    @staticmethod
    def evaluate_edit(original: str, edited: str, instruction: str, intent: str) -> dict:
        """
        Evaluate the quality of an edit.
        
        Returns metrics:
        - preservation_score: How much meaning is preserved (0-1)
        - instruction_adherence: How well the edit follows instruction (0-1)
        - fluency_score: How natural the output sounds (0-1)
        - overall_score: Combined score (0-1)
        """
        # Preservation score based on word overlap
        orig_words = set(original.lower().split())
        edit_words = set(edited.lower().split())
        
        if intent == "summarize":
            # For summarization, we expect less overlap
            preservation = len(orig_words & edit_words) / max(len(orig_words), 1)
            preservation = 1.0 - abs(0.5 - preservation)  # Target ~50% overlap
        else:
            preservation = len(orig_words & edit_words) / max(len(orig_words), 1)
        
        # Instruction adherence (simple heuristic)
        adherence = 1.0
        if "formal" in instruction.lower() and any(c in edited for c in ["don't", "won't", "can't"]):
            adherence -= 0.3
        if "uppercase" in instruction.lower() and edited != edited.upper():
            adherence -= 0.5
        if "lowercase" in instruction.lower() and edited != edited.lower():
            adherence -= 0.5
        
        # Fluency (basic checks)
        fluency = 1.0
        # Penalize very short outputs
        if len(edited) < 10:
            fluency -= 0.3
        # Penalize repeated words
        words = edited.lower().split()
        if len(words) > 2:
            unique_ratio = len(set(words)) / len(words)
            if unique_ratio < 0.5:
                fluency -= 0.3
        
        overall = (preservation * 0.3 + adherence * 0.4 + fluency * 0.3)
        
        return {
            "preservation_score": round(preservation, 2),
            "instruction_adherence": round(max(adherence, 0), 2),
            "fluency_score": round(max(fluency, 0), 2),
            "overall_score": round(max(overall, 0), 2),
        }
    
    @staticmethod
    def run_test_suite() -> list[dict]:
        """Run a suite of test cases."""
        test_cases = [
            {
                "original": "I don't think we can't do this.",
                "instruction": "make more formal",
                "expected_contains": ["do not", "cannot"],
            },
            {
                "original": "This is a very really quite important document.",
                "instruction": "make concise",
                "expected_not_contains": ["very", "really", "quite"],
            },
            {
                "original": "hello world",
                "instruction": "uppercase",
                "expected": "HELLO WORLD",
            },
        ]
        
        results = []
        agent = DocumentEditAgent()
        
        for tc in test_cases:
            # This would need to be run in async context
            # For now, just use rule-based
            result = agent._rule_based_edit(tc["original"], tc["instruction"])
            edited = result["edited_text"]
            
            passed = True
            if "expected" in tc:
                passed = edited == tc["expected"]
            if "expected_contains" in tc:
                passed = all(exp in edited for exp in tc["expected_contains"])
            if "expected_not_contains" in tc:
                passed = all(exp not in edited for exp in tc["expected_not_contains"])
            
            results.append({
                "original": tc["original"],
                "instruction": tc["instruction"],
                "edited": edited,
                "passed": passed,
            })
        
        return results


# Singleton
_agent: DocumentEditAgent | None = None

def get_edit_agent() -> DocumentEditAgent:
    """Get the document edit agent singleton."""
    global _agent
    if _agent is None:
        _agent = DocumentEditAgent()
    return _agent
