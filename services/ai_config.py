"""Centralized AI configuration.

Single source of truth for all AI provider settings.
Reads from environment variables with sensible defaults.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Literal

# Provider priority order
ProviderType = Literal["gemini", "openai", "anthropic", "ollama", "stub"]


@dataclass
class AIProviderConfig:
    """Configuration for a specific AI provider."""
    model_name: str
    api_key: str | None = None
    base_url: str | None = None
    available: bool = False


@dataclass
class AISettings:
    """Centralized AI settings loaded from environment.
    
    Usage:
        settings = get_ai_settings()
        print(settings.primary_provider)  # "gemini"
        print(settings.gemini.model_name)  # "gemini-2.5-flash"
    """
    # Generation settings
    max_output_tokens: int = 2000
    temperature: float = 0.3
    
    # Guardrails
    max_input_length: int = 5000
    max_output_length: int = 10000
    
    # Provider configs
    gemini: AIProviderConfig = field(default_factory=lambda: AIProviderConfig(model_name="gemini-2.5-flash"))
    openai: AIProviderConfig = field(default_factory=lambda: AIProviderConfig(model_name="gpt-4o-mini"))
    anthropic: AIProviderConfig = field(default_factory=lambda: AIProviderConfig(model_name="claude-3-haiku-20240307"))
    ollama: AIProviderConfig = field(default_factory=lambda: AIProviderConfig(model_name="llama3.2", base_url="http://localhost:11434"))
    
    # Determined at load time
    primary_provider: ProviderType = "stub"
    
    # Evals settings
    evals_enabled: bool = True
    evals_log_to_console: bool = True


def _load_settings_from_env() -> AISettings:
    """Load AI settings from environment variables."""
    settings = AISettings()
    
    # Gemini config (PRIMARY for this POC)
    gemini_key = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
    settings.gemini = AIProviderConfig(
        model_name=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
        api_key=gemini_key,
        available=bool(gemini_key),
    )
    
    # OpenAI config (legacy/alternative)
    openai_key = os.getenv("OPENAI_API_KEY")
    settings.openai = AIProviderConfig(
        model_name=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
        api_key=openai_key,
        available=bool(openai_key),
    )
    
    # Anthropic config (legacy/alternative)
    anthropic_key = os.getenv("ANTHROPIC_API_KEY")
    settings.anthropic = AIProviderConfig(
        model_name=os.getenv("ANTHROPIC_MODEL", "claude-3-haiku-20240307"),
        api_key=anthropic_key,
        available=bool(anthropic_key),
    )
    
    # Ollama config (legacy/alternative)
    ollama_host = os.getenv("OLLAMA_HOST", "http://localhost:11434")
    settings.ollama = AIProviderConfig(
        model_name=os.getenv("OLLAMA_MODEL", "llama3.2"),
        base_url=ollama_host,
        available=bool(os.getenv("OLLAMA_HOST")),  # Only if explicitly set
    )
    
    # Generation settings from env
    if os.getenv("AI_MAX_OUTPUT_TOKENS"):
        settings.max_output_tokens = int(os.getenv("AI_MAX_OUTPUT_TOKENS"))
    if os.getenv("AI_TEMPERATURE"):
        settings.temperature = float(os.getenv("AI_TEMPERATURE"))
    
    # Evals settings
    settings.evals_enabled = os.getenv("DISABLE_EVALS", "").lower() not in ("1", "true")
    settings.evals_log_to_console = os.getenv("EVALS_LOG_TO_CONSOLE", "1").lower() in ("1", "true")
    
    # Determine primary provider (Gemini first for this POC)
    if settings.gemini.available:
        settings.primary_provider = "gemini"
    elif settings.openai.available:
        settings.primary_provider = "openai"
    elif settings.anthropic.available:
        settings.primary_provider = "anthropic"
    elif settings.ollama.available:
        settings.primary_provider = "ollama"
    else:
        settings.primary_provider = "stub"
    
    return settings


# Singleton instance
_settings: AISettings | None = None


def get_ai_settings() -> AISettings:
    """Get the AI settings singleton.
    
    Settings are loaded once from environment on first access.
    """
    global _settings
    if _settings is None:
        _settings = _load_settings_from_env()
    return _settings


def reload_ai_settings() -> AISettings:
    """Force reload settings from environment.
    
    Useful for testing or after env changes.
    """
    global _settings
    _settings = _load_settings_from_env()
    return _settings
