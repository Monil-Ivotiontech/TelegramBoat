"""
Message Cleanup Service - Background task for deleting old messages.

Periodically deletes messages from Telegram chats and database.
"""

import asyncio
from typing import Optional

from telegram import Bot

from app.core.config import settings
from app.core.database import db_pool
from app.core.logging import get_logger
from app.bots.base import BotType

logger = get_logger(__name__)


class MessageCleanupService:
    """
    Service for cleaning up old messages from Telegram chats.
    
    Runs as a background task, periodically checking for messages
    in the database and deleting them from both Telegram and the database.
    """

    def __init__(self):
        """Initialize the cleanup service."""
        self._running = False
        self._task: Optional[asyncio.Task] = None
        self._bots: dict = {}
        self._cleanup_interval = 120  # 120 seconds like the old bots
        
    def register_bot(self, bot_type: BotType, bot_instance: Bot) -> None:
        """
        Register a bot instance for message deletion.
        
        Args:
            bot_type: The type of bot
            bot_instance: The Telegram Bot instance
        """
        self._bots[bot_type] = bot_instance
        logger.debug("Bot registered for cleanup", bot_type=bot_type.value)

    async def start(self) -> None:
        """Start the background cleanup task."""
        if self._running:
            logger.warning("Cleanup service already running")
            return
        
        self._running = True
        self._task = asyncio.create_task(self._cleanup_loop())
        logger.info(
            "Message cleanup service started",
            interval=self._cleanup_interval,
        )

    async def stop(self) -> None:
        """Stop the background cleanup task."""
        if not self._running:
            return
        
        self._running = False
        if self._task:
            self._task.cancel()
            try:
                await self._task
            except asyncio.CancelledError:
                pass
            self._task = None
        
        logger.info("Message cleanup service stopped")

    async def _cleanup_loop(self) -> None:
        """Main cleanup loop that runs periodically."""
        while self._running:
            try:
                await self._delete_messages()
            except Exception as e:
                logger.error("Error in cleanup loop", error=str(e))
            
            await asyncio.sleep(self._cleanup_interval)

    async def _delete_messages(self) -> None:
        """Delete messages for all registered bots."""
        # Map bot types to their message tables
        table_mapping = {
            BotType.EMAIL: "email_chat_message_detail",
            BotType.PHONE: "phone_chat_message_detail",
            BotType.CITY: "city_chat_message_detail",
            BotType.ID: "id_chat_message_detail",
            BotType.STATE: "state_chat_message_detail",
            BotType.ZIP: "zip_chat_message_detail",
        }
        
        for bot_type, table_name in table_mapping.items():
            bot = self._bots.get(bot_type)
            if not bot:
                continue
            
            await self._delete_messages_for_table(bot, table_name)

    async def _delete_messages_for_table(
        self, 
        bot: Bot, 
        table_name: str
    ) -> None:
        """
        Delete messages from a specific table.
        
        Args:
            bot: The Telegram Bot instance
            table_name: The database table name
        """
        try:
            # Get messages to delete (limit to avoid processing too many at once)
            # Process 100 messages per cleanup cycle
            query = f"""
                SELECT id, chatid, msgid
                FROM {table_name}
                ORDER BY id ASC
                LIMIT 100
            """
            rows = await db_pool.fetch(query)
            
            logger.info(
                "Cleanup check",
                table=table_name,
                messages_found=len(rows) if rows else 0,
            )
            
            if not rows:
                return
            
            logger.debug(
                "Processing message deletion",
                table=table_name,
                count=len(rows),
            )
            
            deleted_ids = []
            
            for row in rows:
                msg_id = row["id"]
                chat_id = row["chatid"]
                telegram_msg_id = row["msgid"]
                
                try:
                    # Try to delete from Telegram
                    await bot.delete_message(
                        chat_id=chat_id,
                        message_id=telegram_msg_id,
                    )
                except Exception as e:
                    # Message may already be deleted or too old
                    # Still mark for deletion from database
                    logger.debug(
                        "Could not delete message from Telegram",
                        chat_id=chat_id,
                        message_id=telegram_msg_id,
                        error=str(e),
                    )
                
                deleted_ids.append(msg_id)
            
            # Delete from database
            if deleted_ids:
                if len(deleted_ids) == 1:
                    delete_query = f"""
                        DELETE FROM {table_name}
                        WHERE id = $1
                    """
                    await db_pool.execute(delete_query, deleted_ids[0])
                else:
                    # Use ANY for multiple IDs
                    delete_query = f"""
                        DELETE FROM {table_name}
                        WHERE id = ANY($1::int[])
                    """
                    await db_pool.execute(delete_query, deleted_ids)
                
                logger.info(
                    "Messages deleted",
                    table=table_name,
                    count=len(deleted_ids),
                )
                
        except Exception as e:
            logger.error(
                "Failed to delete messages",
                table=table_name,
                error=str(e),
            )


# Global message cleanup service instance
message_cleanup_service = MessageCleanupService()

