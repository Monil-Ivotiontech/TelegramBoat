"""
Bot module containing bot configurations and registry.
"""

from app.bots.registry import BotRegistry, bot_registry
from app.bots.base import BaseBot, BotConfig

__all__ = [
    "BotRegistry",
    "bot_registry",
    "BaseBot",
    "BotConfig",
]

