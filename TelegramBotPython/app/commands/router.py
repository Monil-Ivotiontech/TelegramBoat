"""
Command Router - Routes commands to appropriate handlers.
"""

from typing import Any, Dict, Optional

from app.bots.base import BotConfig
from app.core.logging import get_logger
from app.repositories.command_repo import Command
from app.utils.constants import ReportType

logger = get_logger(__name__)


class CommandRouter:
    """Routes commands to appropriate handlers based on method."""

    async def try_special_commands(
        self,
        message_text: str,
        chat_id: int,
        manager_id: str,
        manager_type: str,
        config: BotConfig,
    ) -> Optional[Dict[str, Any]]:
        """
        Try to handle special commands (symbol, DW, customer ID).

        Args:
            message_text: The message text
            chat_id: Telegram chat ID
            manager_id: Manager ID
            manager_type: Manager type
            config: Bot configuration

        Returns:
            Result dict or None if not a special command
        """
        from app.commands.symbols import SymbolCommands
        from app.commands.deposit_withdrawal import DepositWithdrawalCommands
        from app.commands.customer import CustomerCommands
        from app.services.manager_service import manager_service

        # Get commission type based on dwaccess
        commission_type = await manager_service.get_commission_type(
            manager_id, manager_type, chat_id
        )

        # Try symbol lookup
        symbol_commands = SymbolCommands(config)
        result = await symbol_commands.try_symbol_lookup(
            message_text, chat_id, manager_id, manager_type, commission_type
        )
        if result is not None:
            return result

        # Try deposit/withdrawal
        if message_text.upper().startswith("DW"):
            dw_commands = DepositWithdrawalCommands(config)
            return await dw_commands.get_deposit_withdrawal(
                message_text, chat_id, manager_id, manager_type
            )

        # Try customer ID lookup
        customer_commands = CustomerCommands(config)
        result = await customer_commands.try_customer_lookup(
            message_text, chat_id, manager_id, manager_type
        )
        if result is not None:
            return result

        return None

    async def execute_command(
        self,
        command: Command,
        manager_id: str,
        manager_type: str,
        config: BotConfig,
        chat_id: int,
    ) -> Optional[Dict[str, Any]]:
        """
        Execute a command based on its method.

        Args:
            command: The command to execute
            manager_id: Manager ID
            manager_type: Manager type
            config: Bot configuration
            chat_id: Telegram chat ID

        Returns:
            Result dict or None
        """
        method = command.method
        commission_type = command.commission_type

        logger.info(
            "Executing command",
            command=command.command_name,
            method=method,
            manager_id=manager_id,
            manager_type=manager_type,
        )

        if method == ReportType.M2M:
            from app.commands.m2m import M2MCommands
            m2m = M2MCommands(config)
            return await m2m.get_m2m_report(
                manager_type, manager_id, commission_type
            )

        elif method == ReportType.COM_POS:
            from app.commands.positions import PositionCommands
            positions = PositionCommands(config)
            return await positions.get_com_pos_report(
                manager_type, manager_id, commission_type
            )

        elif method == ReportType.TOTAL_POS:
            from app.commands.positions import PositionCommands
            positions = PositionCommands(config)
            return await positions.get_total_pos_report(
                manager_type, manager_id
            )

        elif method == ReportType.UPDATE_ALL:
            from app.commands.reports import ReportCommands
            reports = ReportCommands(config)
            return await reports.get_update_all_report(
                manager_type, manager_id
            )

        elif method == ReportType.TOP_5:
            from app.commands.reports import ReportCommands
            reports = ReportCommands(config)
            return await reports.get_top_report(
                manager_type, manager_id, commission_type, count=5
            )

        elif method == ReportType.TOP_10:
            from app.commands.reports import ReportCommands
            reports = ReportCommands(config)
            return await reports.get_top_report(
                manager_type, manager_id, commission_type, count=10
            )

        logger.warning("Unknown command method", method=method)
        return None


# Global command router instance
command_router = CommandRouter()

