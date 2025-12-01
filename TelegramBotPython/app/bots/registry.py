"""
Bot Registry - Central management of all bot instances.

Allows easy addition of new bots and unified lifecycle management.
"""

from typing import Dict, Optional, List

from app.bots.base import BaseBot, BotConfig, BotType
from app.core.config import settings
from app.core.logging import get_logger

logger = get_logger(__name__)


class BotRegistry:
    """
    Registry for managing multiple bot instances.
    
    Provides centralized bot lifecycle management and easy
    addition of new bot types.
    """

    def __init__(self):
        """Initialize the registry."""
        self._bots: Dict[BotType, BaseBot] = {}
        self._configs: Dict[BotType, BotConfig] = {}
        
        # Register default bot configurations
        self._register_default_configs()

    def _register_default_configs(self) -> None:
        """Register configurations for all supported bot types."""
        
        # Email Bot Configuration
        self._configs[BotType.EMAIL] = BotConfig(
            bot_type=BotType.EMAIL,
            token=settings.EMAIL_BOT_TOKEN,
            commission_field="Email",  # Corrected to match old bot - uses Email field
            message_table="email_chat_message_detail",
            display_name="📧 Email Bot",
        )

        # Phone Bot Configuration
        self._configs[BotType.PHONE] = BotConfig(
            bot_type=BotType.PHONE,
            token=settings.PHONE_BOT_TOKEN,
            commission_field="Phone",
            message_table="phone_chat_message_detail",
            display_name="📱 Phone Bot",
        )

        # City Bot Configuration
        self._configs[BotType.CITY] = BotConfig(
            bot_type=BotType.CITY,
            token=settings.CITY_BOT_TOKEN,
            commission_field="City",
            message_table="city_chat_message_detail",
            display_name="🏙️ City Bot",
        )

        # ID Bot Configuration
        self._configs[BotType.ID] = BotConfig(
            bot_type=BotType.ID,
            token=settings.ID_BOT_TOKEN,
            commission_field="MT5_ID",
            message_table="id_chat_message_detail",
            display_name="🆔 ID Bot",
        )

        # State Bot Configuration
        self._configs[BotType.STATE] = BotConfig(
            bot_type=BotType.STATE,
            token=settings.STATE_BOT_TOKEN,
            commission_field="State",
            message_table="state_chat_message_detail",
            display_name="🗺️ State Bot",
        )

        # Zip Bot Configuration
        self._configs[BotType.ZIP] = BotConfig(
            bot_type=BotType.ZIP,
            token=settings.ZIP_BOT_TOKEN,
            commission_field="ZipCode",
            message_table="zip_chat_message_detail",
            display_name="📮 Zip Bot",
        )

        logger.info(
            "Bot configurations registered",
            bot_types=[bt.value for bt in self._configs.keys()],
        )

    def get_config(self, bot_type: BotType) -> Optional[BotConfig]:
        """
        Get configuration for a specific bot type.

        Args:
            bot_type: The type of bot

        Returns:
            BotConfig or None if not registered
        """
        return self._configs.get(bot_type)

    def create_bot(self, bot_type: BotType) -> BaseBot:
        """
        Create a bot instance for the given type.

        Args:
            bot_type: The type of bot to create

        Returns:
            BaseBot instance

        Raises:
            ValueError: If bot type is not registered
        """
        config = self.get_config(bot_type)
        if config is None:
            raise ValueError(f"No configuration registered for bot type: {bot_type}")

        bot = BaseBot(config)
        self._bots[bot_type] = bot
        logger.info("Bot created", bot_type=bot_type.value)
        return bot

    def get_bot(self, bot_type: BotType) -> Optional[BaseBot]:
        """
        Get an existing bot instance.

        Args:
            bot_type: The type of bot

        Returns:
            BaseBot instance or None
        """
        return self._bots.get(bot_type)

    async def start_all(self) -> None:
        """Start all registered bots."""
        logger.info("Starting all bots", count=len(self._configs))
        
        for bot_type in self._configs.keys():
            try:
                bot = self.create_bot(bot_type)
                await bot.start()
            except Exception as e:
                logger.error(
                    "Failed to start bot",
                    bot_type=bot_type.value,
                    error=str(e),
                )
                raise

        logger.info("All bots started successfully")

    async def stop_all(self) -> None:
        """Stop all running bots."""
        logger.info("Stopping all bots", count=len(self._bots))
        
        for bot_type, bot in self._bots.items():
            try:
                await bot.stop()
            except Exception as e:
                logger.error(
                    "Error stopping bot",
                    bot_type=bot_type.value,
                    error=str(e),
                )

        self._bots.clear()
        logger.info("All bots stopped")

    def get_running_bots(self) -> List[BotType]:
        """Get list of currently running bot types."""
        return [bt for bt, bot in self._bots.items() if bot.is_running]

    def get_status(self) -> Dict[str, dict]:
        """Get status of all bots for monitoring."""
        status = {}
        for bot_type, config in self._configs.items():
            bot = self._bots.get(bot_type)
            status[bot_type.value] = {
                "configured": True,
                "running": bot.is_running if bot else False,
                "display_name": config.display_name,
            }
        return status


# Global bot registry instance
bot_registry = BotRegistry()

