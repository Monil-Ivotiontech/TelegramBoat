"""
Base Bot Configuration and Abstract Bot Class.

Provides the foundation for all bot types with shared functionality.
"""

from dataclasses import dataclass, field
from typing import Optional, TYPE_CHECKING
from enum import Enum

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

from app.core.logging import get_logger

if TYPE_CHECKING:
    from app.handlers.start import StartHandler
    from app.handlers.message import MessageHandler as MsgHandler
    from app.handlers.callbacks import CallbackHandler

logger = get_logger(__name__)


class BotType(str, Enum):
    """Enumeration of available bot types."""
    EMAIL = "EMAIL"
    PHONE = "PHONE"
    CITY = "CITY"
    STATE = "STATE"
    ZIP = "ZIP"
    ID = "ID"


@dataclass
class BotConfig:
    """
    Configuration for a specific bot type.
    
    This dataclass holds all bot-specific settings that differentiate
    one bot from another (e.g., Email vs Phone bot).
    """
    
    # Bot identification
    bot_type: BotType
    token: str
    
    # Commission field in MT5 user data (Email, Phone, City, etc.)
    commission_field: str
    
    # Database table for storing chat messages
    message_table: str
    
    # Display name for user messages
    display_name: str = ""
    
    # Optional: specific command mappings
    command_prefix: str = ""
    
    def __post_init__(self):
        """Set default display name if not provided."""
        if not self.display_name:
            self.display_name = f"{self.bot_type.value.title()} Bot"


class BaseBot:
    """
    Base class for Telegram bots.
    
    Handles bot lifecycle (start/stop) and handler registration.
    """

    def __init__(self, config: BotConfig):
        """
        Initialize the bot with configuration.

        Args:
            config: Bot-specific configuration
        """
        self.config = config
        self.application: Optional[Application] = None
        self._running = False
        self.logger = get_logger(
            f"Bot.{config.bot_type.value}",
            bot_type=config.bot_type.value,
        )

    async def start(self) -> None:
        """Start the bot and begin polling for updates."""
        if self._running:
            self.logger.warning("Bot already running")
            return

        self.logger.info("Starting bot", display_name=self.config.display_name)

        # Build the application
        self.application = (
            Application.builder()
            .token(self.config.token)
            .build()
        )

        # Register handlers
        self._register_handlers()

        # Initialize and start polling
        await self.application.initialize()
        await self.application.start()
        await self.application.updater.start_polling(
            drop_pending_updates=True,
            allowed_updates=Update.ALL_TYPES,
        )

        self._running = True
        self.logger.info("Bot started successfully", display_name=self.config.display_name)
        
        # Register with message cleanup service
        from app.services.message_cleanup_service import message_cleanup_service
        if self.application and self.application.bot:
            message_cleanup_service.register_bot(
                self.config.bot_type,
                self.application.bot,
            )

    async def stop(self) -> None:
        """Stop the bot gracefully."""
        if not self._running or self.application is None:
            self.logger.warning("Bot not running")
            return

        self.logger.info("Stopping bot", display_name=self.config.display_name)

        # Stop polling and shutdown
        await self.application.updater.stop()
        await self.application.stop()
        await self.application.shutdown()

        self._running = False
        self.logger.info("Bot stopped successfully", display_name=self.config.display_name)

    def _register_handlers(self) -> None:
        """Register all message and command handlers."""
        if self.application is None:
            raise RuntimeError("Application not initialized")

        # Import handlers here to avoid circular imports
        from app.handlers.start import create_start_handler
        from app.handlers.message import create_message_handler
        from app.handlers.callbacks import (
            create_manager_selection_handler,
            create_callback_handler,
        )

        # Register /start command
        self.application.add_handler(
            CommandHandler("start", create_start_handler(self.config))
        )

        # Register manager selection callback (must be before generic callback)
        self.application.add_handler(
            CallbackQueryHandler(
                create_manager_selection_handler(self.config),
                pattern=r"^select_manager_.*"
            )
        )

        # Register generic callback handler
        self.application.add_handler(
            CallbackQueryHandler(create_callback_handler(self.config))
        )

        # Register text message handler (must be last)
        self.application.add_handler(
            MessageHandler(
                filters.TEXT & ~filters.COMMAND,
                create_message_handler(self.config)
            )
        )

        self.logger.debug("Handlers registered")

    @property
    def is_running(self) -> bool:
        """Check if the bot is currently running."""
        return self._running

    @property
    def bot_type(self) -> BotType:
        """Get the bot type."""
        return self.config.bot_type

