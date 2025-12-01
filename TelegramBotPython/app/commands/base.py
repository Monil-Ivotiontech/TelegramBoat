"""
Base Command Class.

Provides common functionality for all command handlers.
"""

from typing import Any, Dict, List, Optional
import pandas as pd
import numpy as np

from app.bots.base import BotConfig
from app.core.logging import get_logger
from app.services.mt5_service import mt5_service
from app.services.manager_service import manager_service
from app.services.pdf_service import pdf_service
from app.utils.formatters import format_currency, safe_float, safe_int
from app.utils.constants import ManagerType, VOLUME_DIVISORS, TradeAction


class BaseCommand:
    """Base class for command handlers with common functionality."""

    def __init__(self, config: BotConfig):
        """
        Initialize the command handler.

        Args:
            config: Bot configuration
        """
        self.config = config
        self.logger = get_logger(self.__class__.__name__)

    async def get_user_data(
        self,
        manager_type: str,
        manager_id: str,
    ) -> tuple[List[Dict], List[str]]:
        """
        Get user data for a manager.

        Returns:
            Tuple of (user_list, group_list)
        """
        if manager_type == ManagerType.MT5:
            groups = await manager_service.get_manager_groups(manager_id, manager_type)
            if not groups:
                return [], []
            
            users = await mt5_service.get_users_by_groups(groups)
            return users, groups
        else:
            logins = await manager_service.get_manager_user_logins(manager_id, manager_type)
            if not logins:
                return [], []
            
            users = await mt5_service.get_users_by_logins(logins)
            return users, []

    def extract_commission(
        self,
        users: List[Dict],
        manager_id: str,
    ) -> pd.DataFrame:
        """
        Extract user data with commission from the commission field.

        Args:
            users: List of user data dicts
            manager_id: Manager ID for default commission

        Returns:
            DataFrame with Login, Name, commission columns
        """
        if not users:
            return pd.DataFrame(columns=["Login", "Name", "commission"])

        df = pd.DataFrame(users)
        
        # Get commission from the configured field (Email, Phone, etc.)
        commission_field = self.config.commission_field
        
        df["commission"] = df.get(commission_field, "")
        df["commission"] = df["commission"].replace(r"", 0)
        df["commission"] = pd.to_numeric(df["commission"], errors="coerce")
        df["commission"] = df["commission"].fillna(0)
        df["commission"] = df["commission"] * 100
        df["commission"] = df["commission"].astype(int)
        df["Login"] = df["Login"].astype(np.int64)

        # Get manager's default commission
        manager_user = df[df["Login"] == int(manager_id)]
        default_commission = 0
        if not manager_user.empty:
            default_commission = manager_user["commission"].values[0]

        # Apply default commission where 0
        df.loc[df["commission"] == 0, "commission"] = default_commission

        return df[["Login", "Name", "commission"]]

    def process_volume(self, df: pd.DataFrame) -> pd.DataFrame:
        """Process volume columns to human-readable values."""
        if "Volume" in df.columns:
            df["Volume"] = df["Volume"].astype(np.int64) / VOLUME_DIVISORS.VOLUME
        if "VolumeExt" in df.columns:
            df["VolumeExt"] = df["VolumeExt"].astype(np.int64) / VOLUME_DIVISORS.VOLUME_EXT
        return df

    def process_action(self, df: pd.DataFrame) -> pd.DataFrame:
        """Convert action codes to BUY/SELL strings."""
        if "Action" in df.columns:
            df["Action"] = df["Action"].astype(int).astype(object)
            df.loc[df["Action"] == TradeAction.MT5_BUY, "Action"] = TradeAction.BUY
            df.loc[df["Action"] == TradeAction.MT5_SELL, "Action"] = TradeAction.SELL
        return df

    def apply_commission_filter(
        self,
        df: pd.DataFrame,
        commission_type: str,
    ) -> pd.DataFrame:
        """
        Apply commission type filter (ours vs theirs).

        Args:
            df: DataFrame with commission column
            commission_type: "ours" or "theirs"

        Returns:
            Modified DataFrame
        """
        if commission_type == "ours" and "commission" in df.columns:
            df["commission"] = 100 - df["commission"]
        return df

    def format_currency_columns(
        self,
        df: pd.DataFrame,
        columns: List[str],
    ) -> pd.DataFrame:
        """Format specified columns as currency."""
        for col in columns:
            if col in df.columns:
                df[col] = df[col].apply(format_currency)
        return df

    def add_summary_row(
        self,
        df: pd.DataFrame,
        sum_columns: List[str],
        label_column: str = "Name",
        label_value: str = "Total",
    ) -> pd.DataFrame:
        """Add a summary row to the DataFrame."""
        if df.empty:
            return df

        summary = {}
        for col in df.columns:
            if col in sum_columns:
                summary[col] = df[col].sum()
            elif col == label_column:
                summary[col] = label_value
            else:
                summary[col] = ""

        return pd.concat([df, pd.DataFrame([summary])], ignore_index=True)

    async def generate_pdf_result(
        self,
        pdf_data,
        filename: str,
    ) -> Dict[str, Any]:
        """Create a PDF result dictionary."""
        return {
            "type": "pdf",
            "data": pdf_data,
            "filename": filename,
        }

    def create_text_result(
        self,
        text: str,
        reply_markup: Optional[Any] = None,
    ) -> Dict[str, Any]:
        """Create a text result dictionary."""
        result = {"type": "text", "data": text}
        if reply_markup:
            result["reply_markup"] = reply_markup
        return result

