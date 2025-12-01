"""
M2M Command Handler.

Generates Mark-to-Market reports.
"""

from typing import Any, Dict, Optional
import pandas as pd
import numpy as np

from app.commands.base import BaseCommand
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.utils.formatters import format_currency
from app.utils.constants import ManagerType


class M2MCommands(BaseCommand):
    """Handler for M2M (Mark-to-Market) reports."""

    async def get_m2m_report(
        self,
        manager_type: str,
        manager_id: str,
        commission_type: Optional[str],
    ) -> Dict[str, Any]:
        """
        Generate M2M report.

        Args:
            manager_type: Type of manager ("venus" or "mt5")
            manager_id: Manager ID
            commission_type: Commission type ("ours" or "theirs")

        Returns:
            PDF result dictionary
        """
        self.logger.info(
            "Generating M2M report",
            manager_id=manager_id,
            manager_type=manager_type,
            commission_type=commission_type,
        )

        empty_df = pd.DataFrame(
            columns=["Login", "Name", "Per %", "Floating P&L", "Book P&L", 
                     "M2M", "Partner", "Net Amount"]
        )

        try:
            # Get user data
            users, groups = await self.get_user_data(manager_type, manager_id)
            if not users:
                self.logger.warning("No users found for manager", manager_id=manager_id)
                pdf = await pdf_service.generate_m2m_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "M2M.pdf")

            # Extract commission data
            user_df = self.extract_commission(users, manager_id)

            # Get trade status
            if manager_type == ManagerType.MT5:
                trade_data = await mt5_service.get_trade_accounts_by_groups(groups)
            else:
                logins = user_df["Login"].tolist()
                trade_data = await mt5_service.get_trade_accounts_by_logins(logins)

            if not trade_data:
                self.logger.warning("No trade data found", manager_id=manager_id)
                pdf = await pdf_service.generate_m2m_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "M2M.pdf")

            # Create trade DataFrame
            trade_df = pd.DataFrame(trade_data)
            trade_df = trade_df[["Login", "Credit", "Equity", "Profit", "Balance"]]
            trade_df["Login"] = trade_df["Login"].astype(np.int64)
            trade_df["Credit"] = trade_df["Credit"].astype(float)
            trade_df["Equity"] = trade_df["Equity"].astype(float)
            trade_df["Profit"] = trade_df["Profit"].astype(float)
            trade_df["Balance"] = trade_df["Balance"].astype(float)

            # Merge user and trade data
            result_df = pd.merge(user_df, trade_df, on=["Login"])

            if result_df.empty:
                pdf = await pdf_service.generate_m2m_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "M2M.pdf")

            # Calculate M2M values
            result_df["M2M"] = result_df["Equity"] - result_df["Credit"]
            result_df["Partner"] = result_df["M2M"] * (result_df["commission"] / 100)
            result_df["Net Amount"] = result_df["M2M"] * ((100 - result_df["commission"]) / 100)

            # Round values
            result_df["M2M"] = result_df["M2M"].round(0).astype(int)
            result_df["Partner"] = result_df["Partner"].round(0).astype(int)
            result_df["Net Amount"] = result_df["Net Amount"].round(0).astype(int)

            # Apply commission type
            if commission_type == "ours":
                result_df["commission"] = 100 - result_df["commission"]

            # Calculate totals before filtering
            total_floating_pl = result_df["Profit"].sum()
            total_book_pl = result_df["Balance"].sum()
            total_m2m = result_df["M2M"].sum()
            total_partner = result_df["Partner"].sum()
            total_net_amount = result_df["Net Amount"].sum()

            # Format commission as percentage
            result_df["commission"] = result_df["commission"].astype(str) + "%"

            # Filter out zero M2M rows for display
            result_df = result_df[result_df["M2M"] != 0]

            # Handle all-zero M2M case (from old Email bot logic)
            # Even if all individual M2M values are 0, we should show the total row
            if result_df.empty:
                self.logger.info(
                    "All M2M values are zero - returning summary only",
                    manager_id=manager_id,
                )
                # Create empty result with just the summary row
                summary_only = pd.DataFrame([{
                    "Login": "",
                    "Name": "Total",
                    "Per %": "",
                    "Floating P&L": round(total_floating_pl, 2),
                    "Book P&L": round(total_book_pl, 2),
                    "M2M": format_currency(total_m2m),
                    "Partner": format_currency(total_partner),
                    "Net Amount": format_currency(total_net_amount),
                }])
                pdf = await pdf_service.generate_m2m_pdf(summary_only)
                return await self.generate_pdf_result(pdf, "M2M.pdf")

            # Rename columns
            result_df = result_df.rename(columns={
                "commission": "Per %",
                "Profit": "Floating P&L",
                "Balance": "Book P&L",
            })

            result_df = result_df.sort_values(by=["Login"])

            # Select final columns
            result_df = result_df[[
                "Login", "Name", "Per %", "Floating P&L", "Book P&L",
                "M2M", "Partner", "Net Amount"
            ]]

            # Add summary row
            summary = {
                "Login": "",
                "Name": "Total",
                "Per %": "",
                "Floating P&L": round(total_floating_pl, 2),
                "Book P&L": round(total_book_pl, 2),
                "M2M": total_m2m,
                "Partner": total_partner,
                "Net Amount": total_net_amount,
            }
            result_df = pd.concat([result_df, pd.DataFrame([summary])], ignore_index=True)

            # Round P&L columns
            result_df["Floating P&L"] = result_df["Floating P&L"].round(2)
            result_df["Book P&L"] = result_df["Book P&L"].round(2)

            # Format currency columns
            result_df["M2M"] = result_df["M2M"].apply(format_currency)
            result_df["Partner"] = result_df["Partner"].apply(format_currency)
            result_df["Net Amount"] = result_df["Net Amount"].apply(format_currency)

            self.logger.info(
                "M2M report generated",
                manager_id=manager_id,
                rows=len(result_df),
            )

            pdf = await pdf_service.generate_m2m_pdf(result_df)
            return await self.generate_pdf_result(pdf, "M2M.pdf")

        except Exception as e:
            self.logger.error(
                "M2M report generation failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            pdf = await pdf_service.generate_m2m_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "M2M.pdf")

