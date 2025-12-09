from __future__ import annotations

import os
from dotenv import load_dotenv

# Load .env file BEFORE any other imports that might need env vars
load_dotenv()

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from api.routes.documents import router as documents_router
from api.routes.evals import router as evals_router
from api.routes.spreadsheets import router as spreadsheets_router
from middleware.rate_limit import RateLimitMiddleware, RateLimitConfig
from services.db import init_db


app = FastAPI(title="Document Digital Copy POC")

# Rate limiting (can be disabled in dev with DISABLE_RATE_LIMIT=1)
if not os.getenv("DISABLE_RATE_LIMIT"):
    rate_config = RateLimitConfig(
        requests_per_minute=60,
        requests_per_hour=1000,
        ai_requests_per_minute=20,  # More lenient for AI endpoints in dev
        ai_requests_per_hour=200,
        burst_limit=15,
    )
    app.add_middleware(RateLimitMiddleware, config=rate_config)

# Allow any origin in local dev / POC mode.
# This should be tightened for production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Ensure DB schema exists
init_db()

app.include_router(documents_router)
app.include_router(spreadsheets_router)
app.include_router(evals_router)


@app.get("/")
async def root():
    return {"status": "ok"}
