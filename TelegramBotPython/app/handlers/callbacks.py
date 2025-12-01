"""
Callback Query Handlers.

Handles inline button callbacks.
"""

from typing import Callable

from telegram import Update
from telegram.ext import ContextTypes

from app.bots.base import BotConfig
from app.core.logging import get_logger
from app.core.exceptions import VenusBotError
from app.services.manager_service import manager_service
from app.repositories.message_repo import message_repo
from app.handlers.start import create_start_handler

logger = get_logger(__name__)


def create_manager_selection_handler(config: BotConfig) -> Callable:
    """
    Create a handler for manager selection callbacks.

    Args:
        config: Bot configuration

    Returns:
        Async handler function
    """

    async def manager_selection_handler(
        update: Update, context: ContextTypes.DEFAULT_TYPE
    ) -> None:
        """Handle manager selection callback."""
        query = update.callback_query
        chat_id = update.effective_chat.id

        try:
            # Parse callback data: select_manager_{id}_{type}
            data = query.data.split("_")
            if len(data) < 4:
                await query.answer("Invalid selection")
                return

            manager_id = data[2]
            manager_type = data[3]

            # Validate manager
            managers = await manager_service.get_authorized_managers(
                chat_id, config.bot_type
            )
            
            valid = any(
                m.manager_id == manager_id and m.manager_type == manager_type
                for m in managers
            )

            if not valid:
                await query.answer("Manager no longer available")
                await query.edit_message_text("Manager selection invalid. Please use /start.")
                return

            # Save selection
            context.user_data["selected_manager"] = manager_id
            context.user_data["selected_managertype"] = manager_type

            logger.info(
                "Manager selected",
                chat_id=chat_id,
                manager_id=manager_id,
                manager_type=manager_type,
            )

            await query.answer()
            await query.edit_message_text(
                f"Selected Manager {manager_id} ({manager_type.upper()}). Loading..."
            )

            # Trigger start command to show menu
            start_handler = create_start_handler(config)
            await start_handler(update, context)

        except Exception as e:
            logger.error(
                "Manager selection failed",
                chat_id=chat_id,
                error=str(e),
                exc_info=True,
            )
            await query.answer("Error selecting manager")

    return manager_selection_handler


def create_callback_handler(config: BotConfig) -> Callable:
    """
    Create a generic callback handler for other inline buttons.

    Args:
        config: Bot configuration

    Returns:
        Async handler function
    """

    async def callback_handler(
        update: Update, context: ContextTypes.DEFAULT_TYPE
    ) -> None:
        """Handle generic callback queries (customer actions, etc.)."""
        query = update.callback_query
        chat_id = update.effective_chat.id
        callback_data = query.data

        logger.debug(
            "Callback received",
            chat_id=chat_id,
            data=callback_data,
            bot_type=config.bot_type.value,
        )

        try:
            await query.answer()

            # Route to appropriate handler based on callback data
            result = await process_callback(
                callback_data=callback_data,
                chat_id=chat_id,
                context=context,
                config=config,
            )

            if result is None:
                return

            # Send result
            if result.get("type") == "pdf":
                msg = await update.effective_chat.send_document(
                    document=result["data"],
                    filename=result.get("filename", "report.pdf"),
                )
                await message_repo.save_message(
                    chat_id, msg.message_id, config.message_table
                )
            elif result.get("type") == "text":
                msg = await update.effective_chat.send_message(result["data"])
                await message_repo.save_message(
                    chat_id, msg.message_id, config.message_table
                )

        except VenusBotError as e:
            logger.warning(
                "Callback error",
                chat_id=chat_id,
                error=str(e),
            )
            msg = await update.effective_chat.send_message(e.user_message)
            await message_repo.save_message(chat_id, msg.message_id, config.message_table)
            
        except Exception as e:
            logger.error(
                "Callback handler failed",
                chat_id=chat_id,
                error=str(e),
                exc_info=True,
            )
            msg = await update.effective_chat.send_message(
                "Something went wrong. Please try again."
            )
            await message_repo.save_message(chat_id, msg.message_id, config.message_table)

    return callback_handler


async def process_callback(
    callback_data: str,
    chat_id: int,
    context: ContextTypes.DEFAULT_TYPE,
    config: BotConfig,
) -> dict | None:
    """
    Process a callback query and return the result.

    Args:
        callback_data: The callback data string
        chat_id: Telegram chat ID
        context: Bot context
        config: Bot configuration

    Returns:
        Result dict or None
    """
    from app.commands.customer import CustomerCommands

    customer_commands = CustomerCommands(config)

    # Position callback
    if callback_data.startswith("Position of "):
        customer_id = int(callback_data.replace("Position of ", ""))
        return await customer_commands.get_customer_position(customer_id)

    # Today's Trades callback
    if callback_data.startswith("Today's Trades of "):
        customer_id = int(callback_data.replace("Today's Trades of ", ""))
        return await customer_commands.get_customer_trades(customer_id, "today")

    # This Week's Trades callback
    if callback_data.startswith("This Week's Trades of "):
        customer_id = int(callback_data.replace("This Week's Trades of ", ""))
        return await customer_commands.get_customer_trades(customer_id, "week")

    # Today's Bill callback
    if callback_data.startswith("Today's Bill of "):
        customer_id = int(callback_data.replace("Today's Bill of ", ""))
        return await customer_commands.get_customer_bill(customer_id, "today")

    # This Week's Bill callback
    if callback_data.startswith("This Week's Bill of "):
        customer_id = int(callback_data.replace("This Week's Bill of ", ""))
        return await customer_commands.get_customer_bill(customer_id, "week")

    logger.debug("Unknown callback", data=callback_data)
    return None

