"""
Command Repository - Database access for commands and mappings.
"""

from dataclasses import dataclass
from typing import List, Optional

from app.core.database import db_pool
from app.core.logging import get_logger
from app.core.exceptions import CommandNotFoundError

logger = get_logger(__name__)


@dataclass
class Command:
    """Represents a bot command."""
    
    command_id: int
    command_name: str
    command_type: str
    command_desc: Optional[str]
    commission_type: Optional[str]
    method: Optional[str]
    seq_no: int


@dataclass
class CommandMapping:
    """Represents a command-manager mapping."""
    
    command_id: int
    manager_id: int
    manager_type: str
    method: str


class CommandRepository:
    """Repository for command-related database operations."""

    async def get_commands_for_manager(
        self,
        manager_id: str,
        manager_type: str,
    ) -> List[Command]:
        """
        Get all commands available to a manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager ("venus" or "mt5")

        Returns:
            List of Command objects
        """
        query = """
            SELECT 
                c.id as command_id,
                c.commandname as command_name,
                c.command_type,
                c.commanddesc as command_desc,
                c.commission_type,
                cm.method,
                c.command_seq_no as seq_no
            FROM venus_command c
            JOIN venus_command_mapping cm ON c.id = cm.commandid
            WHERE cm.managerid = $1 
              AND cm.managertype = $2
              AND c.utilflag = 'active'
              AND cm.utilflag = 'active'
            ORDER BY c.command_seq_no
        """
        
        rows = await db_pool.fetch(query, int(manager_id), manager_type)
        
        return [
            Command(
                command_id=row["command_id"],
                command_name=row["command_name"],
                command_type=row["command_type"],
                command_desc=row["command_desc"],
                commission_type=row["commission_type"],
                method=row["method"],
                seq_no=row["seq_no"],
            )
            for row in rows
        ]

    async def get_command_names_for_manager(
        self,
        manager_id: str,
        manager_type: str,
    ) -> List[str]:
        """
        Get command names available to a manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager

        Returns:
            List of command names
        """
        commands = await self.get_commands_for_manager(manager_id, manager_type)
        return [cmd.command_name for cmd in commands]

    async def get_command_by_name(self, command_name: str) -> Optional[Command]:
        """
        Get a command by its name.

        Args:
            command_name: Name of the command

        Returns:
            Command object or None
        """
        query = """
            SELECT 
                id as command_id,
                commandname as command_name,
                command_type,
                commanddesc as command_desc,
                commission_type,
                command_seq_no as seq_no
            FROM venus_command
            WHERE commandname = $1 AND utilflag = 'active'
        """
        
        row = await db_pool.fetchrow(query, command_name)
        
        if not row:
            return None
        
        return Command(
            command_id=row["command_id"],
            command_name=row["command_name"],
            command_type=row["command_type"],
            command_desc=row["command_desc"],
            commission_type=row["commission_type"],
            method=None,
            seq_no=row["seq_no"],
        )

    async def get_command_mapping(
        self,
        command_id: int,
        manager_id: str,
    ) -> Optional[CommandMapping]:
        """
        Get command mapping for a specific command and manager.

        Args:
            command_id: Command ID
            manager_id: Manager ID

        Returns:
            CommandMapping object or None
        """
        query = """
            SELECT commandid, managerid, managertype, method
            FROM venus_command_mapping
            WHERE commandid = $1 AND managerid = $2 AND utilflag = 'active'
        """
        
        row = await db_pool.fetchrow(query, command_id, int(manager_id))
        
        if not row:
            return None
        
        return CommandMapping(
            command_id=row["commandid"],
            manager_id=row["managerid"],
            manager_type=row["managertype"],
            method=row["method"],
        )

    async def validate_command_for_manager(
        self,
        command_name: str,
        manager_id: str,
    ) -> Command:
        """
        Validate that a command exists and is available to a manager.

        Args:
            command_name: Name of the command
            manager_id: Manager ID

        Returns:
            Command object with method

        Raises:
            CommandNotFoundError: If command not found or not available
        """
        command = await self.get_command_by_name(command_name)
        
        if not command:
            raise CommandNotFoundError(
                command=command_name,
                user_message="Command not found.",
            )
        
        mapping = await self.get_command_mapping(command.command_id, manager_id)
        
        if not mapping:
            raise CommandNotFoundError(
                command=command_name,
                message=f"Command '{command_name}' not available for manager {manager_id}",
                user_message="Command not found.",
            )
        
        command.method = mapping.method
        return command


# Global command repository instance
command_repo = CommandRepository()

