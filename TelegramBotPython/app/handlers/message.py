"""
Message Handler.

Routes text messages to appropriate command handlers.
"""

from typing import Callable

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from app.bots.base import BotConfig
from app.core.logging import get_logger
from app.core.exceptions import VenusBotError, CommandNotFoundError
from app.services.manager_service import manager_service
from app.repositories.command_repo import command_repo
from app.repositories.message_repo import message_repo

logger = get_logger(__name__)


def create_message_handler(config: BotConfig) -> Callable:
    """
    Create a message handler for a specific bot configuration.

    Args:
        config: Bot configuration

    Returns:
        Async handler function
    """

    async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Handle text messages."""
        chat_id = update.effective_chat.id
        message_text = update.message.text
        message_id = update.message.message_id

        logger.debug(
            "Message received",
            chat_id=chat_id,
            text=message_text[:50],
            bot_type=config.bot_type.value,
        )

        # Save incoming message
        await message_repo.save_message(chat_id, message_id, config.message_table)

        try:
            # Get authorized managers
            managers = await manager_service.get_authorized_managers(
                chat_id, config.bot_type
            )

            if not managers:
                msg = await update.effective_chat.send_message(
                    "You are not authorized to use this bot."
                )
                await message_repo.save_message(chat_id, msg.message_id, config.message_table)
                return

            # Check if manager selection is needed
            selected_manager = context.user_data.get("selected_manager")
            selected_type = context.user_data.get("selected_managertype")

            if len(managers) > 1 and not selected_manager:
                # Store the original message for later processing
                context.user_data["original_message"] = message_text
                
                keyboard = [
                    [InlineKeyboardButton(
                        f"Manager {m.manager_id} ({m.manager_type.upper()})",
                        callback_data=f"select_manager_{m.manager_id}_{m.manager_type}"
                    )]
                    for m in managers
                ]
                reply_markup = InlineKeyboardMarkup(keyboard)
                
                msg = await update.effective_chat.send_message(
                    "Multiple managers found. Please select one to proceed:",
                    reply_markup=reply_markup,
                )
                await message_repo.save_message(chat_id, msg.message_id, config.message_table)
                return

            # Use selected manager or default to first
            if not selected_manager:
                selected_manager = managers[0].manager_id
                selected_type = managers[0].manager_type
                context.user_data["selected_manager"] = selected_manager
                context.user_data["selected_managertype"] = selected_type

            # Validate manager still exists
            valid_manager = next(
                (m for m in managers 
                 if m.manager_id == selected_manager and m.manager_type == selected_type),
                None
            )
            
            if not valid_manager:
                msg = await update.effective_chat.send_message(
                    "Invalid manager selection. Please use /start to select again."
                )
                await message_repo.save_message(chat_id, msg.message_id, config.message_table)
                return

            # Check command access
            if valid_manager.command_access == "DISABLE":
                msg = await update.effective_chat.send_message(
                    "We are facing some technical issues. Please try again later."
                )
                await message_repo.save_message(chat_id, msg.message_id, config.message_table)
                return

            # Route to command processor
            await process_command(
                update=update,
                context=context,
                config=config,
                manager_id=selected_manager,
                manager_type=selected_type,
                message_text=message_text,
            )

        except VenusBotError as e:
            logger.warning(
                "Bot error",
                chat_id=chat_id,
                error=str(e),
                user_message=e.user_message,
            )
            msg = await update.effective_chat.send_message(e.user_message)
            await message_repo.save_message(chat_id, msg.message_id, config.message_table)
            
        except Exception as e:
            logger.error(
                "Message handler failed",
                chat_id=chat_id,
                error=str(e),
                exc_info=True,
            )
            msg = await update.effective_chat.send_message(
                "Something went wrong. Please try again."
            )
            await message_repo.save_message(chat_id, msg.message_id, config.message_table)

    return message_handler


async def process_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    config: BotConfig,
    manager_id: str,
    manager_type: str,
    message_text: str,
) -> None:
    """
    Process a command message.

    Routes to the appropriate command handler based on the message.
    """
    from app.commands import command_router

    chat_id = update.effective_chat.id

    # Check for special commands first (symbol lookup, DW, customer ID)
    result = await command_router.try_special_commands(
        message_text=message_text,
        chat_id=chat_id,
        manager_id=manager_id,
        manager_type=manager_type,
        config=config,
    )

    if result is not None:
        # Special command handled
        if result.get("type") == "pdf":
            msg = await update.message.reply_document(
                document=result["data"],
                filename=result.get("filename", "report.pdf"),
            )
        elif result.get("type") == "text":
            if result.get("reply_markup"):
                msg = await update.effective_chat.send_message(
                    result["data"],
                    reply_markup=result["reply_markup"],
                )
            else:
                msg = await update.effective_chat.send_message(result["data"])
        elif result.get("type") == "excel":
            msg = await update.message.reply_document(
                document=result["data"],
                filename=result.get("filename", "report.xlsx"),
            )
        else:
            return
            
        await message_repo.save_message(chat_id, msg.message_id, config.message_table)
        return

    # Try standard command lookup
    try:
        command = await command_repo.validate_command_for_manager(
            message_text, manager_id
        )
    except CommandNotFoundError:
        msg = await update.effective_chat.send_message("Command not found.")
        await message_repo.save_message(chat_id, msg.message_id, config.message_table)
        return

    # Get command mapping to determine actual manager type
    mapping = await command_repo.get_command_mapping(command.command_id, manager_id)
    actual_manager_type = mapping.manager_type if mapping else manager_type

    # Execute the command
    result = await command_router.execute_command(
        command=command,
        manager_id=manager_id,
        manager_type=actual_manager_type,
        config=config,
        chat_id=chat_id,
    )

    if result is None:
        msg = await update.effective_chat.send_message(
            "Under Development. Coming Soon."
        )
        await message_repo.save_message(chat_id, msg.message_id, config.message_table)
        return

    # Send result
    if result.get("type") == "pdf":
        msg = await update.message.reply_document(
            document=result["data"],
            filename=result.get("filename", "report.pdf"),
        )
    elif result.get("type") == "text":
        msg = await update.effective_chat.send_message(result["data"])
    else:
        return
        
    await message_repo.save_message(chat_id, msg.message_id, config.message_table)

