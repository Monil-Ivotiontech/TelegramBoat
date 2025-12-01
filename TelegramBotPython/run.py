#!/usr/bin/env python3
"""
VenusBot Entry Point

This script starts the FastAPI application with all configured bots.
Usage:
    poetry run python run.py
    # or
    python run.py
"""

import asyncio
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

import uvicorn
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


def main():
    """Main entry point for the VenusBot application."""
    from app.core.config import settings
    
    print("=" * 60)
    print("🚀 VenusBot Unified System")
    print("=" * 60)
    print(f"📧 Email Bot: Enabled")
    print(f"📱 Phone Bot: Enabled")
    print(f"🌐 Server: http://{settings.HOST}:{settings.PORT}")
    print(f"📊 Environment: {settings.APP_ENV}")
    print("=" * 60)
    
    uvicorn.run(
        "app.main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG,
        log_level=settings.LOG_LEVEL.lower(),
    )


if __name__ == "__main__":
    main()

