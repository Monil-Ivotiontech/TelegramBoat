"""
Application Configuration using Pydantic Settings.

All configuration is loaded from environment variables with validation.
"""

from functools import lru_cache
from typing import Literal

from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore",
    )

    # -------------------------------------------------------------------------
    # Bot Tokens
    # -------------------------------------------------------------------------
    EMAIL_BOT_TOKEN: str = Field(..., description="Telegram bot token for Email bot")
    PHONE_BOT_TOKEN: str = Field(..., description="Telegram bot token for Phone bot")
    CITY_BOT_TOKEN: str = Field(default="", description="Telegram bot token for City bot")
    ID_BOT_TOKEN: str = Field(default="", description="Telegram bot token for ID bot")
    STATE_BOT_TOKEN: str = Field(default="", description="Telegram bot token for State bot")
    ZIP_BOT_TOKEN: str = Field(default="", description="Telegram bot token for Zip bot")

    # -------------------------------------------------------------------------
    # Database Configuration
    # -------------------------------------------------------------------------
    DB_HOST: str = Field(default="localhost", description="Database host")
    DB_PORT: int = Field(default=5432, description="Database port")
    DB_USER: str = Field(default="postgres", description="Database user")
    DB_PASSWORD: str = Field(..., description="Database password")
    DB_NAME: str = Field(default="serverlotusfxauth", description="Database name")

    # Connection Pool
    DB_POOL_MIN_SIZE: int = Field(default=2, ge=1, le=10)
    DB_POOL_MAX_SIZE: int = Field(default=10, ge=2, le=50)

    # -------------------------------------------------------------------------
    # MT5 Configuration
    # -------------------------------------------------------------------------
    MT5_SERVER: str = Field(default="api.lotuus.co", description="MT5 API server")
    MT5_LOGIN: int = Field(default=1039, description="MT5 login ID")
    MT5_PASSWORD: str = Field(..., description="MT5 password")

    # -------------------------------------------------------------------------
    # Application Settings
    # -------------------------------------------------------------------------
    APP_ENV: Literal["development", "production", "testing"] = Field(default="production")
    LOG_LEVEL: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = Field(default="INFO")
    DEBUG: bool = Field(default=False)

    # Server
    HOST: str = Field(default="0.0.0.0")
    PORT: int = Field(default=8000)

    # -------------------------------------------------------------------------
    # Computed Properties
    # -------------------------------------------------------------------------
    @property
    def database_url(self) -> str:
        """Get the async database URL for asyncpg."""
        return f"postgresql://{self.DB_USER}:{self.DB_PASSWORD}@{self.DB_HOST}:{self.DB_PORT}/{self.DB_NAME}"

    @property
    def database_url_async(self) -> str:
        """Get the async database URL for asyncpg."""
        return f"postgresql://{self.DB_USER}:{self.DB_PASSWORD}@{self.DB_HOST}:{self.DB_PORT}/{self.DB_NAME}"

    @field_validator("DB_POOL_MAX_SIZE")
    @classmethod
    def validate_pool_size(cls, v: int, info) -> int:
        """Ensure max pool size is greater than min pool size."""
        min_size = info.data.get("DB_POOL_MIN_SIZE", 2)
        if v < min_size:
            raise ValueError(f"DB_POOL_MAX_SIZE ({v}) must be >= DB_POOL_MIN_SIZE ({min_size})")
        return v


@lru_cache()
def get_settings() -> Settings:
    """Get cached settings instance."""
    return Settings()


# Global settings instance
settings = get_settings()

