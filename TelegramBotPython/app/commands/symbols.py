"""
Symbol Command Handlers.

Handles symbol-specific position lookups.
"""

from typing import Any, Dict, List, Optional
import pandas as pd
import numpy as np

from app.commands.base import BaseCommand
from app.core.database import db_pool
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.repositories.symbol_repo import symbol_repo
from app.utils.constants import ManagerType


class SymbolCommands(BaseCommand):
    """Handler for symbol-related commands."""

    async def try_symbol_lookup(
        self,
        message_text: str,
        chat_id: int,
        manager_id: str,
        manager_type: str,
        commission_type: str = "ours",
    ) -> Optional[Dict[str, Any]]:
        """
        Try to process message as symbol lookup.

        Returns None if not a valid symbol.
        """
        # Get known symbols
        categories = await symbol_repo.get_symbol_categories()
        all_symbols = set()
        for symbols in categories.values():
            all_symbols.update(symbols)

        # Check if message matches a symbol prefix
        message_upper = message_text.upper().strip()
        
        # Match symbols that START with the query (like old version)
        matching_symbols = [
            s for s in all_symbols 
            if s.upper().startswith(message_upper)
        ]
        
        if not matching_symbols:
            return None

        # Check dwaccess for commission type
        actual_commission_type = await self._get_commission_type(
            chat_id, manager_id, manager_type
        )

        # Get positions for the symbol
        return await self.get_symbol_position(
            message_text,
            manager_id,
            manager_type,
            matching_symbols,
            actual_commission_type,
        )

    async def _get_commission_type(
        self,
        chat_id: int,
        manager_id: str,
        manager_type: str,
    ) -> str:
        """Determine commission type based on dwaccess permission."""
        if manager_type == ManagerType.MT5:
            query = """
                SELECT cd.dwaccess
                FROM mt5_manager_chatid_detail cd
                JOIN mt5_manager_chatid c ON cd."Login" = c."Login"
                WHERE cd."Login" = $1 AND c.chatid = $2
                  AND c.utilflag = 'active' AND cd.utilflag = 'active'
            """
            row = await db_pool.fetchrow(query, int(manager_id), chat_id)
            if row and row.get("dwaccess") == "DISABLE":
                return "partner"
        return "ours"

    async def get_symbol_position(
        self,
        symbol_query: str,
        manager_id: str,
        manager_type: str,
        matching_symbols: List[str],
        commission_type: str = "ours",
    ) -> Dict[str, Any]:
        """
        Get position details for specific symbols with per-user breakdown.
        
        Returns report with columns: Code, Name, Per., Type, Total, Partner, NetVolume
        """
        self.logger.info(
            "Getting symbol position",
            symbol_query=symbol_query,
            manager_id=manager_id,
            commission_type=commission_type,
        )

        empty_df = pd.DataFrame(
            columns=["Code", "Name", "Per.", "Type", "Total", "Partner", "NetVolume"]
        )

        try:
            # Get users with commission data
            users, groups = await self.get_user_data(manager_type, manager_id)
            if not users:
                return self.create_text_result("No data found.")

            user_df = self.extract_commission(users, manager_id)

            # Get positions
            if manager_type == ManagerType.MT5:
                positions = await mt5_service.get_positions_by_groups(groups)
            else:
                logins = user_df["Login"].tolist()
                positions = await mt5_service.get_positions_by_logins(logins)

            if not positions:
                return self.create_text_result("No positions found.")

            position_df = pd.DataFrame(positions)
            position_df = position_df[["Login", "Symbol", "Action", "Volume"]]
            position_df["Login"] = position_df["Login"].astype(np.int64)

            # Merge with user data
            merged_df = pd.merge(user_df, position_df, on=["Login"])

            if merged_df.empty:
                return self.create_text_result(f"No positions found for {symbol_query}.")

            # Process data
            merged_df["Volume"] = merged_df["Volume"].astype(np.int64) / 10000
            merged_df["Action"] = merged_df["Action"].astype(int)
            merged_df.loc[merged_df["Action"] == 0, "Action"] = "BUY"
            merged_df.loc[merged_df["Action"] == 1, "Action"] = "SELL"

            # Make SELL volumes negative for calculations
            merged_df.loc[merged_df["Action"] == "SELL", "Volume"] *= -1

            # Calculate Partner and NetVolume based on commission
            merged_df["Partner"] = (merged_df["Volume"] * merged_df["commission"]) / 100
            merged_df["NetVolume"] = (merged_df["Volume"] * (100 - merged_df["commission"])) / 100

            # Clean symbol names
            merged_df["Symbol"] = merged_df["Symbol"].str.split("/").str[0]

            # Filter out zero volumes
            merged_df = merged_df[merged_df["Volume"] != 0]

            # Apply commission type display
            if commission_type == "partner":
                merged_df["commission"] = 100 - merged_df["commission"]

            # Rename columns
            merged_df = merged_df.rename(columns={
                "Login": "Code",
                "commission": "Per.",
                "Action": "Type",
                "Volume": "Total"
            })

            merged_df = merged_df[[
                "Symbol", "Code", "Name", "Per.", "Type", "Total", "Partner", "NetVolume"
            ]]

            # Replace inf values
            merged_df.replace([np.inf, -np.inf], 0, inplace=True)
            merged_df = merged_df.fillna(0)

            # Round values
            merged_df["Partner"] = merged_df["Partner"].round(0).astype(int)
            merged_df["NetVolume"] = merged_df["NetVolume"].round(0).astype(int)

            # Format percentage
            merged_df["Per."] = merged_df["Per."].astype(str) + "%"

            # Filter by matching symbols and create per-symbol reports
            result_data = []
            for symbol in matching_symbols:
                symbol_df = merged_df[merged_df["Symbol"] == symbol].copy()
                
                if symbol_df.empty:
                    continue

                # Remove Symbol column for display
                symbol_df = symbol_df[[
                    "Code", "Name", "Per.", "Type", "Total", "Partner", "NetVolume"
                ]]

                # Add summary row
                total_row = {
                    "Code": "",
                    "Name": "Total",
                    "Per.": "",
                    "Type": "SELL" if symbol_df["Total"].sum() <= 0 else "BUY",
                    "Total": symbol_df["Total"].sum(),
                    "Partner": symbol_df["Partner"].sum(),
                    "NetVolume": symbol_df["NetVolume"].sum(),
                }

                # Make individual row values absolute for display
                symbol_df["Total"] = symbol_df["Total"].abs()
                symbol_df["Partner"] = symbol_df["Partner"].abs()
                symbol_df["NetVolume"] = symbol_df["NetVolume"].abs()
                
                symbol_df = symbol_df.sort_values(by=["Code"])
                symbol_df = pd.concat(
                    [symbol_df, pd.DataFrame([total_row])],
                    ignore_index=True
                )

                result_data.append({
                    "symbol": symbol,
                    "df": symbol_df,
                })

            if not result_data:
                return self.create_text_result(f"No positions found for {symbol_query}.")

            pdf = await pdf_service.generate_symbol_position_pdf(result_data)
            return await self.generate_pdf_result(pdf, f"Symbol_{symbol_query}.pdf")

        except Exception as e:
            self.logger.error(
                "Symbol lookup failed",
                symbol_query=symbol_query,
                error=str(e),
                exc_info=True,
            )
            return self.create_text_result("Symbol lookup failed. Please try again.")

