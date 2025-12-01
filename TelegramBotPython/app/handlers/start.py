"""
Start Command Handler.

Handles the /start command for bot initialization.
"""

from typing import Callable

from telegram import Update, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from app.bots.base import BotConfig
from app.core.logging import get_logger
from app.services.manager_service import manager_service
from app.repositories.command_repo import command_repo
from app.repositories.message_repo import message_repo

logger = get_logger(__name__)


def create_start_handler(config: BotConfig) -> Callable:
    """
    Create a start command handler for a specific bot configuration.

    Args:
        config: Bot configuration

    Returns:
        Async handler function
    """

    async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        """Handle the /start command."""
        chat_id = update.effective_chat.id
        
        logger.info(
            "Start command received",
            chat_id=chat_id,
            bot_type=config.bot_type.value,
        )

        try:
            # Reset manager selection on fresh /start
            if update.message and update.message.text == "/start":
                context.user_data.pop("selected_manager", None)
                context.user_data.pop("selected_managertype", None)

            # Get authorized managers
            managers = await manager_service.get_authorized_managers(
                chat_id, config.bot_type
            )

            if not managers:
                await update.effective_chat.send_message(
                    "You are not authorized to use this bot."
                )
                return

            # Check if manager is already selected
            selected_manager = context.user_data.get("selected_manager")
            selected_type = context.user_data.get("selected_managertype")

            if selected_manager and selected_type:
                # Validate the selected manager still exists
                valid = any(
                    m.manager_id == selected_manager and m.manager_type == selected_type
                    for m in managers
                )
                if not valid:
                    context.user_data.pop("selected_manager", None)
                    context.user_data.pop("selected_managertype", None)
                    selected_manager = None
                    selected_type = None

            # If no manager selected, prompt or auto-select
            if not selected_manager or not selected_type:
                if len(managers) > 1:
                    # Multiple managers - ask user to select
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
                    await message_repo.save_message(
                        chat_id, msg.message_id, config.message_table
                    )
                    return
                else:
                    # Single manager - auto-select
                    selected_manager = managers[0].manager_id
                    selected_type = managers[0].manager_type
                    context.user_data["selected_manager"] = selected_manager
                    context.user_data["selected_managertype"] = selected_type

            # Get available commands for the manager
            command_names = await command_repo.get_command_names_for_manager(
                selected_manager, selected_type
            )

            if not command_names:
                await update.effective_chat.send_message("No commands available.")
                return

            # Create keyboard with commands
            buttons = [[KeyboardButton(cmd)] for cmd in command_names]
            reply_markup = ReplyKeyboardMarkup(buttons, resize_keyboard=True)

            msg = await update.effective_chat.send_message(
                f"Welcome! Using Manager {selected_manager} ({selected_type.upper()})\n"
                f"Select a command from the menu below:",
                reply_markup=reply_markup,
            )
            
            await message_repo.save_message(
                chat_id, msg.message_id, config.message_table
            )

        except Exception as e:
            logger.error(
                "Start command failed",
                chat_id=chat_id,
                error=str(e),
                exc_info=True,
            )
            await update.effective_chat.send_message(
                "Something went wrong. Please try again."
            )

    return start_handler

