"""
Async Database Connection Pool using asyncpg.

Provides a singleton connection pool for efficient database access.
"""

from contextlib import asynccontextmanager
from typing import Any, AsyncGenerator, Optional

import asyncpg
from asyncpg import Pool, Connection, Record

from app.core.config import settings
from app.core.logging import get_logger
from app.core.exceptions import DatabaseError

logger = get_logger(__name__)


class DatabasePool:
    """
    Async database connection pool manager.
    
    Implements singleton pattern to ensure only one pool exists.
    """

    _instance: Optional["DatabasePool"] = None
    _pool: Optional[Pool] = None

    def __new__(cls) -> "DatabasePool":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    async def connect(self) -> None:
        """Initialize the connection pool."""
        if self._pool is not None:
            logger.warning("Database pool already initialized")
            return

        try:
            self._pool = await asyncpg.create_pool(
                host=settings.DB_HOST,
                port=settings.DB_PORT,
                user=settings.DB_USER,
                password=settings.DB_PASSWORD,
                database=settings.DB_NAME,
                min_size=settings.DB_POOL_MIN_SIZE,
                max_size=settings.DB_POOL_MAX_SIZE,
                command_timeout=30,
                max_inactive_connection_lifetime=300,
            )
            logger.info(
                "Database pool initialized",
                host=settings.DB_HOST,
                database=settings.DB_NAME,
                pool_size=f"{settings.DB_POOL_MIN_SIZE}-{settings.DB_POOL_MAX_SIZE}",
            )
        except Exception as e:
            logger.error("Failed to initialize database pool", error=str(e))
            raise DatabaseError(f"Failed to connect to database: {e}")

    async def disconnect(self) -> None:
        """Close the connection pool."""
        if self._pool is not None:
            await self._pool.close()
            self._pool = None
            logger.info("Database pool closed")

    @asynccontextmanager
    async def acquire(self) -> AsyncGenerator[Connection, None]:
        """
        Acquire a connection from the pool.

        Usage:
            async with db_pool.acquire() as conn:
                result = await conn.fetch("SELECT * FROM table")
        """
        if self._pool is None:
            raise DatabaseError("Database pool not initialized")

        async with self._pool.acquire() as connection:
            yield connection

    async def fetch(self, query: str, *args: Any) -> list[Record]:
        """
        Execute a query and fetch all results.

        Args:
            query: SQL query with $1, $2, etc. placeholders
            *args: Query parameters

        Returns:
            List of records
        """
        async with self.acquire() as conn:
            return await conn.fetch(query, *args)

    async def fetchrow(self, query: str, *args: Any) -> Optional[Record]:
        """
        Execute a query and fetch a single row.

        Args:
            query: SQL query with $1, $2, etc. placeholders
            *args: Query parameters

        Returns:
            Single record or None
        """
        async with self.acquire() as conn:
            return await conn.fetchrow(query, *args)

    async def fetchval(self, query: str, *args: Any) -> Any:
        """
        Execute a query and fetch a single value.

        Args:
            query: SQL query with $1, $2, etc. placeholders
            *args: Query parameters

        Returns:
            Single value
        """
        async with self.acquire() as conn:
            return await conn.fetchval(query, *args)

    async def execute(self, query: str, *args: Any) -> str:
        """
        Execute a query without returning results.

        Args:
            query: SQL query with $1, $2, etc. placeholders
            *args: Query parameters

        Returns:
            Command status string
        """
        async with self.acquire() as conn:
            return await conn.execute(query, *args)

    async def executemany(self, query: str, args: list[tuple]) -> None:
        """
        Execute a query multiple times with different parameters.

        Args:
            query: SQL query with $1, $2, etc. placeholders
            args: List of parameter tuples
        """
        async with self.acquire() as conn:
            await conn.executemany(query, args)

    @property
    def pool(self) -> Optional[Pool]:
        """Get the underlying connection pool."""
        return self._pool

    def get_pool_status(self) -> dict:
        """Get current pool status for monitoring."""
        if self._pool is None:
            return {"status": "not_initialized"}

        return {
            "status": "active",
            "size": self._pool.get_size(),
            "free_size": self._pool.get_idle_size(),
            "min_size": self._pool.get_min_size(),
            "max_size": self._pool.get_max_size(),
        }


# Global database pool instance
db_pool = DatabasePool()


def get_db_pool() -> DatabasePool:
    """Get the database pool instance (for dependency injection)."""
    return db_pool

