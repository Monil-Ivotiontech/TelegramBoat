"""
Manager Service - Business logic for manager authorization and validation.

Handles Venus and MT5 manager lookups and authorization checks.
"""

from dataclasses import dataclass
from typing import List, Optional, Tuple

from app.core.database import db_pool
from app.core.logging import get_logger
from app.core.exceptions import AuthorizationError, ManagerNotFoundError
from app.bots.base import BotConfig, BotType
from app.utils.constants import ManagerType

logger = get_logger(__name__)


@dataclass
class Manager:
    """Represents a manager (Venus or MT5)."""
    
    manager_id: str
    manager_type: str  # "venus" or "mt5"
    chat_id: int
    commission: Optional[int] = None
    command_access: str = "ENABLE"
    dw_access: str = "ENABLE"  # DW access permission
    exp_date: Optional[str] = None


class ManagerService:
    """
    Service for manager-related operations.
    
    Handles authorization, manager lookup, and validation.
    """

    async def get_authorized_managers(
        self,
        chat_id: int,
        bot_type: BotType,
    ) -> List[Manager]:
        """
        Get all managers authorized for a chat ID and bot type.

        Args:
            chat_id: Telegram chat ID
            bot_type: Type of bot (EMAIL, PHONE, etc.)

        Returns:
            List of authorized Manager objects
        """
        managers = []

        # Get Venus managers
        venus_managers = await self._get_venus_managers(chat_id, bot_type)
        managers.extend(venus_managers)

        # Get MT5 managers
        mt5_managers = await self._get_mt5_managers(chat_id, bot_type)
        managers.extend(mt5_managers)

        logger.debug(
            "Found authorized managers",
            chat_id=chat_id,
            bot_type=bot_type.value,
            count=len(managers),
        )

        return managers

    async def _get_venus_managers(
        self,
        chat_id: int,
        bot_type: BotType,
    ) -> List[Manager]:
        """Get Venus managers for a chat ID."""
        query = """
            SELECT vm.id as manager_id, vm.chatid, vm.commandaccess, vm.expdate
            FROM venus_manager vm
            JOIN venus_manager_bot_access vmba ON vm.id = vmba.managerid
            WHERE vm.chatid = $1 
              AND vmba.bottype = $2 
              AND vm.utilflag = 'active'
              AND vmba.utilflag = 'active'
        """
        
        rows = await db_pool.fetch(query, chat_id, bot_type.value)
        
        return [
            Manager(
                manager_id=str(row["manager_id"]),
                manager_type=ManagerType.VENUS,
                chat_id=row["chatid"],
                command_access=row["commandaccess"] or "ENABLE",
                exp_date=row["expdate"],
            )
            for row in rows
        ]

    async def _get_mt5_managers(
        self,
        chat_id: int,
        bot_type: BotType,
    ) -> List[Manager]:
        """Get MT5 managers for a chat ID."""
        query = """
            SELECT cd."Login" as manager_id, c.chatid, cd.commandaccess, cd.dwaccess
            FROM mt5_manager_chatid_detail cd
            JOIN mt5_manager_chatid c ON cd."Login" = c."Login"
            JOIN mt5_manager_bot_access ba ON cd."Login" = ba."Login"
            WHERE c.chatid = $1 
              AND ba.bottype = $2
              AND c.utilflag = 'active'
              AND cd.utilflag = 'active'
              AND ba.utilflag = 'active'
        """
        
        rows = await db_pool.fetch(query, chat_id, bot_type.value)
        
        return [
            Manager(
                manager_id=str(row["manager_id"]),
                manager_type=ManagerType.MT5,
                chat_id=row["chatid"],
                command_access=row["commandaccess"] or "ENABLE",
                dw_access=row["dwaccess"] or "ENABLE",
            )
            for row in rows
        ]

    async def validate_manager(
        self,
        manager_id: str,
        manager_type: str,
        chat_id: int,
        bot_type: BotType,
    ) -> Manager:
        """
        Validate that a manager exists and is authorized.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager ("venus" or "mt5")
            chat_id: Telegram chat ID
            bot_type: Type of bot

        Returns:
            Validated Manager object

        Raises:
            ManagerNotFoundError: If manager not found or not authorized
        """
        managers = await self.get_authorized_managers(chat_id, bot_type)
        
        for manager in managers:
            if manager.manager_id == manager_id and manager.manager_type == manager_type:
                return manager
        
        raise ManagerNotFoundError(
            manager_id=manager_id,
            message=f"Manager {manager_id} ({manager_type}) not authorized for chat {chat_id}",
        )

    async def get_manager_user_logins(
        self,
        manager_id: str,
        manager_type: str,
    ) -> List[int]:
        """
        Get all user logins mapped to a manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager

        Returns:
            List of user login IDs
        """
        if manager_type == ManagerType.MT5:
            # For MT5 managers, get users from groups
            from app.services.mt5_service import mt5_service
            
            manager_detail = await mt5_service.get_manager_detail(int(manager_id))
            if not manager_detail:
                return []
            
            groups = [grp["Group"] for grp in manager_detail[0].get("Groups", [])]
            if not groups:
                return []
            
            users = await mt5_service.get_users_by_groups(groups)
            return [user["Login"] for user in users]
        else:
            # For Venus managers, get from mapping table
            query = """
                SELECT loginid 
                FROM venus_manager_mapping 
                WHERE managerid = $1 AND utilflag = 'active'
            """
            rows = await db_pool.fetch(query, int(manager_id))
            return [row["loginid"] for row in rows]

    async def get_manager_groups(
        self,
        manager_id: str,
        manager_type: str,
    ) -> List[str]:
        """
        Get groups assigned to an MT5 manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager

        Returns:
            List of group names
        """
        if manager_type != ManagerType.MT5:
            return []
        
        from app.services.mt5_service import mt5_service
        
        manager_detail = await mt5_service.get_manager_detail(int(manager_id))
        if not manager_detail:
            return []
        
        return [grp["Group"] for grp in manager_detail[0].get("Groups", [])]

    async def check_command_access(
        self,
        manager_id: str,
        manager_type: str,
        chat_id: int,
    ) -> bool:
        """
        Check if command access is enabled for a manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager
            chat_id: Telegram chat ID

        Returns:
            True if access is enabled
        """
        if manager_type == ManagerType.MT5:
            query = """
                SELECT cd.commandaccess
                FROM mt5_manager_chatid_detail cd
                JOIN mt5_manager_chatid c ON cd."Login" = c."Login"
                WHERE cd."Login" = $1 AND c.chatid = $2
                  AND c.utilflag = 'active' AND cd.utilflag = 'active'
            """
            row = await db_pool.fetchrow(query, int(manager_id), chat_id)
        else:
            query = """
                SELECT commandaccess
                FROM venus_manager
                WHERE id = $1 AND chatid = $2 AND utilflag = 'active'
            """
            row = await db_pool.fetchrow(query, int(manager_id), chat_id)
        
        if not row:
            return False
        
        return row["commandaccess"] != "DISABLE"

    async def check_dw_access(
        self,
        manager_id: str,
        manager_type: str,
        chat_id: int,
    ) -> bool:
        """
        Check if DW (deposit/withdrawal) access is enabled for a manager.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager
            chat_id: Telegram chat ID

        Returns:
            True if DW access is enabled
        """
        if manager_type == ManagerType.MT5:
            query = """
                SELECT cd.dwaccess
                FROM mt5_manager_chatid_detail cd
                JOIN mt5_manager_chatid c ON cd."Login" = c."Login"
                WHERE cd."Login" = $1 AND c.chatid = $2
                  AND c.utilflag = 'active' AND cd.utilflag = 'active'
            """
            row = await db_pool.fetchrow(query, int(manager_id), chat_id)
        else:
            # Venus managers don't have DW access restrictions in old code
            return True
        
        if not row:
            return False
        
        return row.get("dwaccess") != "DISABLE"

    async def get_commission_type(
        self,
        manager_id: str,
        manager_type: str,
        chat_id: int,
    ) -> str:
        """
        Get commission type based on DW access.

        Args:
            manager_id: Manager ID
            manager_type: Type of manager
            chat_id: Telegram chat ID

        Returns:
            "ours" if has full access, "partner" otherwise
        """
        has_access = await self.check_dw_access(manager_id, manager_type, chat_id)
        return "ours" if has_access else "partner"


# Global manager service instance
manager_service = ManagerService()

