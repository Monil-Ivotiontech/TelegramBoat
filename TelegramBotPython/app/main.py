"""
FastAPI Application with Lifespan Events.

This is the main entry point that starts all bots when the server starts
and gracefully shuts them down when the server stops.
"""

from contextlib import asynccontextmanager
from typing import AsyncGenerator

from fastapi import FastAPI
from fastapi.responses import JSONResponse

from app.core.config import settings
from app.core.database import db_pool
from app.core.logging import setup_logging, get_logger
from app.services.mt5_service import mt5_service
from app.services.message_cleanup_service import message_cleanup_service
from app.bots.registry import bot_registry

# Setup logging
setup_logging(log_level=settings.LOG_LEVEL, json_format=settings.APP_ENV == "production")

logger = get_logger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncGenerator:
    """
    FastAPI lifespan context manager.
    
    Handles startup and shutdown of all services and bots.
    """
    # ========== STARTUP ==========
    logger.info("=" * 60)
    logger.info("🚀 Starting VenusBot Unified System")
    logger.info("=" * 60)

    try:
        # Initialize database connection pool
        logger.info("📦 Initializing database connection pool...")
        await db_pool.connect()
        logger.info("✅ Database pool connected")

        # Initialize MT5 client
        logger.info("📡 Initializing MT5 API client...")
        await mt5_service.initialize()
        logger.info("✅ MT5 client initialized")

        # Start all bots
        logger.info("🤖 Starting Telegram bots...")
        await bot_registry.start_all()
        logger.info("✅ All bots started successfully!")

        # # Start message cleanup service
        # logger.info("🧹 Starting message cleanup service...")
        # await message_cleanup_service.start()
        # logger.info("✅ Message cleanup service started")

        logger.info("=" * 60)
        logger.info("🎉 VenusBot System is ready!")
        logger.info(f"📊 Environment: {settings.APP_ENV}")
        logger.info(f"🌐 Health check: http://{settings.HOST}:{settings.PORT}/health")
        logger.info("=" * 60)

    except Exception as e:
        logger.error(f"❌ Startup failed: {e}", exc_info=True)
        raise

    yield  # Application is running

    # ========== SHUTDOWN ==========
    logger.info("=" * 60)
    logger.info("🛑 Shutting down VenusBot System...")

    try:
        # Stop message cleanup service
        logger.info("🧹 Stopping message cleanup service...")
        await message_cleanup_service.stop()
        logger.info("✅ Message cleanup service stopped")

        # Stop all bots
        logger.info("🤖 Stopping Telegram bots...")
        await bot_registry.stop_all()
        logger.info("✅ All bots stopped")

        # Close MT5 client
        logger.info("📡 Closing MT5 client...")
        await mt5_service.close()
        logger.info("✅ MT5 client closed")

        # Close database pool
        logger.info("📦 Closing database pool...")
        await db_pool.disconnect()
        logger.info("✅ Database pool closed")

    except Exception as e:
        logger.error(f"Error during shutdown: {e}", exc_info=True)

    logger.info("✅ Graceful shutdown complete!")
    logger.info("=" * 60)


# Create FastAPI application
app = FastAPI(
    title="VenusBot",
    description="Unified Telegram Bot System for Venus Trading Platform",
    version="1.0.0",
    lifespan=lifespan,
)


# ========== Health & Status Endpoints ==========

@app.get("/health")
async def health_check():
    """
    Health check endpoint.
    
    Returns the current status of all services.
    """
    return JSONResponse({
        "status": "healthy",
        "service": "VenusBot",
        "version": "1.0.0",
        "environment": settings.APP_ENV,
    })


@app.get("/status")
async def get_status():
    """
    Get detailed status of all components.
    
    Returns status of database, MT5, and all bots.
    """
    return JSONResponse({
        "status": "running",
        "database": db_pool.get_pool_status(),
        "bots": bot_registry.get_status(),
        "mt5": {
            "server": settings.MT5_SERVER,
            "initialized": mt5_service._client is not None,
        },
    })


@app.get("/bots")
async def get_bots():
    """
    Get list of running bots.
    """
    return JSONResponse({
        "running_bots": [bt.value for bt in bot_registry.get_running_bots()],
        "all_bots": bot_registry.get_status(),
    })


# ========== Error Handlers ==========

@app.exception_handler(Exception)
async def global_exception_handler(request, exc):
    """Global exception handler."""
    logger.error(f"Unhandled exception: {exc}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"error": "Internal server error"},
    )

