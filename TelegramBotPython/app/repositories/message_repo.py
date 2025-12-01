"""
Message Repository - Database access for chat message tracking.
"""

from app.core.database import db_pool
from app.core.logging import get_logger

logger = get_logger(__name__)


class MessageRepository:
    """Repository for message tracking database operations."""

    async def save_message(
        self,
        chat_id: int,
        message_id: int,
        table_name: str,
    ) -> None:
        """
        Save a message detail to the database.

        Args:
            chat_id: Telegram chat ID
            message_id: Telegram message ID
            table_name: Name of the table to save to
        """
        # Sanitize table name to prevent SQL injection
        # Only allow known table names
        allowed_tables = {
            "email_chat_message_detail",
            "phone_chat_message_detail",
            "city_chat_message_detail",
            "state_chat_message_detail",
            "zip_chat_message_detail",
            "id_chat_message_detail",
        }
        
        if table_name not in allowed_tables:
            logger.error("Invalid table name", table_name=table_name)
            return
        
        try:
            query = f"""
                INSERT INTO {table_name} (chatid, msgid)
                VALUES ($1, $2)
            """
            await db_pool.execute(query, chat_id, message_id)
            logger.debug(
                "Message saved",
                chat_id=chat_id,
                message_id=message_id,
                table=table_name,
            )
        except Exception as e:
            logger.error(
                "Failed to save message",
                chat_id=chat_id,
                message_id=message_id,
                error=str(e),
            )

    async def get_messages_for_deletion(
        self,
        table_name: str,
        limit: int = 100,
    ) -> list:
        """
        Get messages that can be deleted (older messages).

        Args:
            table_name: Name of the table
            limit: Maximum number of messages to return

        Returns:
            List of (chat_id, message_id) tuples
        """
        allowed_tables = {
            "email_chat_message_detail",
            "phone_chat_message_detail",
            "city_chat_message_detail",
            "state_chat_message_detail",
            "zip_chat_message_detail",
            "id_chat_message_detail",
        }
        
        if table_name not in allowed_tables:
            logger.error("Invalid table name", table_name=table_name)
            return []
        
        query = f"""
            SELECT chatid, msgid
            FROM {table_name}
            ORDER BY id ASC
            LIMIT $1
        """
        
        rows = await db_pool.fetch(query, limit)
        return [(row["chatid"], row["msgid"]) for row in rows]

    async def delete_message_record(
        self,
        chat_id: int,
        message_id: int,
        table_name: str,
    ) -> None:
        """
        Delete a message record from the database.

        Args:
            chat_id: Telegram chat ID
            message_id: Telegram message ID
            table_name: Name of the table
        """
        allowed_tables = {
            "email_chat_message_detail",
            "phone_chat_message_detail",
            "city_chat_message_detail",
            "state_chat_message_detail",
            "zip_chat_message_detail",
            "id_chat_message_detail",
        }
        
        if table_name not in allowed_tables:
            return
        
        query = f"""
            DELETE FROM {table_name}
            WHERE chatid = $1 AND msgid = $2
        """
        
        await db_pool.execute(query, chat_id, message_id)


# Global message repository instance
message_repo = MessageRepository()

