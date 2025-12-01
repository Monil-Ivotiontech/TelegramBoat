"""
Core module containing configuration, database, logging, and exceptions.
"""

from app.core.config import settings
from app.core.database import DatabasePool, get_db_pool
from app.core.logging import get_logger, setup_logging
from app.core.exceptions import (
    VenusBotError,
    AuthorizationError,
    CommandNotFoundError,
    MT5ConnectionError,
    DatabaseError,
)

__all__ = [
    "settings",
    "DatabasePool",
    "get_db_pool",
    "get_logger",
    "setup_logging",
    "VenusBotError",
    "AuthorizationError",
    "CommandNotFoundError",
    "MT5ConnectionError",
    "DatabaseError",
]

