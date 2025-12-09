"""Rate limiting middleware for FastAPI."""
from __future__ import annotations

import time
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Callable, Dict

from fastapi import Request, Response
from fastapi.responses import JSONResponse
from starlette.middleware.base import BaseHTTPMiddleware


@dataclass
class RateLimitConfig:
    """Configuration for rate limiting."""
    requests_per_minute: int = 60
    requests_per_hour: int = 1000
    ai_requests_per_minute: int = 10  # Stricter limit for AI endpoints
    ai_requests_per_hour: int = 100
    burst_limit: int = 10  # Max requests in 1 second


@dataclass
class ClientState:
    """Track request counts for a client."""
    minute_requests: list = field(default_factory=list)
    hour_requests: list = field(default_factory=list)
    second_requests: list = field(default_factory=list)
    
    def cleanup(self, now: float):
        """Remove old timestamps."""
        minute_ago = now - 60
        hour_ago = now - 3600
        second_ago = now - 1
        
        self.minute_requests = [t for t in self.minute_requests if t > minute_ago]
        self.hour_requests = [t for t in self.hour_requests if t > hour_ago]
        self.second_requests = [t for t in self.second_requests if t > second_ago]
    
    def record_request(self, now: float):
        """Record a new request."""
        self.minute_requests.append(now)
        self.hour_requests.append(now)
        self.second_requests.append(now)


class RateLimitMiddleware(BaseHTTPMiddleware):
    """Rate limiting middleware with per-client tracking."""
    
    def __init__(self, app, config: RateLimitConfig = None):
        super().__init__(app)
        self.config = config or RateLimitConfig()
        self.clients: Dict[str, ClientState] = defaultdict(ClientState)
    
    def _get_client_id(self, request: Request) -> str:
        """Get a unique identifier for the client."""
        # Use X-Forwarded-For if behind a proxy, otherwise use client host
        forwarded = request.headers.get("X-Forwarded-For")
        if forwarded:
            return forwarded.split(",")[0].strip()
        return request.client.host if request.client else "unknown"
    
    def _is_ai_endpoint(self, path: str) -> bool:
        """Check if this is an AI endpoint (stricter limits)."""
        return "/ai-edit" in path
    
    async def dispatch(self, request: Request, call_next: Callable) -> Response:
        """Process the request with rate limiting."""
        client_id = self._get_client_id(request)
        now = time.time()
        
        # Get or create client state
        state = self.clients[client_id]
        state.cleanup(now)
        
        # Determine limits based on endpoint
        is_ai = self._is_ai_endpoint(request.url.path)
        minute_limit = self.config.ai_requests_per_minute if is_ai else self.config.requests_per_minute
        hour_limit = self.config.ai_requests_per_hour if is_ai else self.config.requests_per_hour
        
        # Check limits
        if len(state.second_requests) >= self.config.burst_limit:
            return JSONResponse(
                status_code=429,
                content={
                    "detail": "Rate limit exceeded: too many requests per second",
                    "retry_after": 1,
                },
                headers={"Retry-After": "1"},
            )
        
        if len(state.minute_requests) >= minute_limit:
            retry_after = 60 - (now - state.minute_requests[0])
            return JSONResponse(
                status_code=429,
                content={
                    "detail": f"Rate limit exceeded: {minute_limit} requests per minute",
                    "retry_after": int(retry_after),
                },
                headers={"Retry-After": str(int(retry_after))},
            )
        
        if len(state.hour_requests) >= hour_limit:
            retry_after = 3600 - (now - state.hour_requests[0])
            return JSONResponse(
                status_code=429,
                content={
                    "detail": f"Rate limit exceeded: {hour_limit} requests per hour",
                    "retry_after": int(retry_after),
                },
                headers={"Retry-After": str(int(retry_after))},
            )
        
        # Record this request
        state.record_request(now)
        
        # Add rate limit headers to response
        response = await call_next(request)
        response.headers["X-RateLimit-Limit"] = str(minute_limit)
        response.headers["X-RateLimit-Remaining"] = str(minute_limit - len(state.minute_requests))
        response.headers["X-RateLimit-Reset"] = str(int(now + 60))
        
        return response


# Convenience function to create configured middleware
def create_rate_limit_middleware(
    requests_per_minute: int = 60,
    ai_requests_per_minute: int = 10,
) -> type:
    """Create a rate limit middleware class with custom config."""
    config = RateLimitConfig(
        requests_per_minute=requests_per_minute,
        ai_requests_per_minute=ai_requests_per_minute,
    )
    
    class ConfiguredRateLimitMiddleware(RateLimitMiddleware):
        def __init__(self, app):
            super().__init__(app, config)
    
    return ConfiguredRateLimitMiddleware
