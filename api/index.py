"""
Vercel serverless entrypoint.

Vercel's @vercel/python runtime detects the ASGI `app` object exported here
and serves it. All routes are rewritten to this function via vercel.json, so
FastAPI handles the full request path (frontend at `/`, API under `/api/...`,
docs at `/docs`).
"""

import sys
from pathlib import Path

# Ensure the project root (where main.py and golmi.xlsx live) is importable.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from main import app  # noqa: E402

# Expose the ASGI app for the Vercel Python runtime.
__all__ = ["app"]
