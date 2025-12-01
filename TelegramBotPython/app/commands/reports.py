"""
Report Command Handlers.

Handles UPDATE ALL, TOP 5, TOP 10 reports.
"""

from typing import Any, Dict, Optional, List
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

from app.commands.base import BaseCommand
from app.core.database import db_pool
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.utils.formatters import format_currency
from app.utils.constants import ManagerType


class ReportCommands(BaseCommand):
    """Handler for various report commands."""

    async def get_update_all_report(
        self,
        manager_type: str,
        manager_id: str,
    ) -> Dict[str, Any]:
        """Generate UPDATE ALL report."""
        self.logger.info(
            "Generating UPDATE ALL report",
            manager_id=manager_id,
            manager_type=manager_type,
        )

        empty_df = pd.DataFrame(
            columns=["MASTER", "NAME", "Profit & Loss"]
        )

        try:
            if manager_type == ManagerType.MT5:
                return await self._get_update_all_mt5(manager_id)
            else:
                return await self._get_update_all_venus(manager_id)

        except Exception as e:
            self.logger.error(
                "UPDATE ALL report failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

    async def _get_update_all_mt5(self, manager_id: str) -> Dict[str, Any]:
        """Generate UPDATE ALL report for MT5 managers."""
        empty_df = pd.DataFrame(columns=["MASTER", "NAME", "Profit & Loss"])

        # Get manager groups
        from app.services.manager_service import manager_service
        groups = await manager_service.get_manager_groups(manager_id, ManagerType.MT5)
        if not groups:
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

        # Get users and extract commission
        users = await mt5_service.get_users_by_groups(groups)
        if not users:
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

        user_df = self.extract_commission(users, manager_id)

        # Get trade accounts
        accounts = await mt5_service.get_trade_accounts_by_groups(groups)
        if not accounts:
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

        account_df = pd.DataFrame(accounts)
        account_df = account_df[["Login", "Credit", "Equity"]]
        account_df["Login"] = account_df["Login"].astype(np.int64)
        account_df["Credit"] = account_df["Credit"].astype(float)
        account_df["Equity"] = account_df["Equity"].astype(float)

        # Merge
        result_df = pd.merge(user_df, account_df, on=["Login"])

        if result_df.empty:
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

        # Calculate M2M and Net Amount
        result_df["M2M"] = result_df["Equity"] - result_df["Credit"]
        result_df["Net Amount"] = result_df["M2M"] * ((100 - result_df["commission"]) / 100)
        result_df["M2M"] = result_df["M2M"].round(2)
        result_df["Net Amount"] = result_df["Net Amount"].round(2)

        result_df = result_df[["Login", "Name", "Net Amount"]]
        result_df = result_df[result_df["Net Amount"] != 0]

        result_df = result_df.rename(columns={
            "Login": "MASTER",
            "Name": "NAME",
            "Net Amount": "Profit & Loss"
        })
        result_df = result_df.sort_values(by=["MASTER"])

        # Add summary
        total_pl = result_df["Profit & Loss"].sum()
        summary = {"MASTER": "", "NAME": "Total", "Profit & Loss": round(total_pl, 2)}
        result_df = pd.concat([result_df, pd.DataFrame([summary])], ignore_index=True)

        result_df["Profit & Loss"] = result_df["Profit & Loss"].apply(format_currency)

        pdf = await pdf_service.generate_update_all_pdf(result_df)
        return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

    async def _get_update_all_venus(self, manager_id: str) -> Dict[str, Any]:
        """
        Generate UPDATE ALL report for Venus managers.
        
        Uses complex group-based aggregation across multiple MT5 managers.
        """
        empty_df = pd.DataFrame(columns=["MASTER", "NAME", "Profit & Loss"])

        try:
            # Get Venus manager group mappings
            manager_group_query = """
                SELECT id, managerid, groupid
                FROM venus_manager_group_mapping
                WHERE managerid = $1 AND utilflag = 'active'
            """
            manager_group_rows = await db_pool.fetch(manager_group_query, int(manager_id))
            
            if not manager_group_rows:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            group_ids = [row["groupid"] for row in manager_group_rows]

            # Get Venus groups
            groups_query = """
                SELECT id, master, name, "desc"
                FROM venus_group
                WHERE id = ANY($1::int[]) AND utilflag = 'active'
            """
            groups_rows = await db_pool.fetch(groups_query, group_ids)
            
            if not groups_rows:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            # Get group mappings (login IDs)
            group_mapping_query = """
                SELECT id, groupid, loginid
                FROM venus_group_mapping
                WHERE groupid = ANY($1::int[]) AND utilflag = 'active'
            """
            group_mapping_rows = await db_pool.fetch(group_mapping_query, group_ids)
            
            if not group_mapping_rows:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            # Get all MT5 managers
            all_managers = await mt5_service.get_manager_detail(0)  # This needs all managers
            # Alternative: Get total and fetch
            total_manager_data = await mt5_service._request("/api/manager/total")
            total_managers = total_manager_data.get("total", 0) if total_manager_data else 0
            
            if total_managers == 0:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            # Fetch all managers
            all_managers = await mt5_service._request(
                "/api/manager/next",
                {"index": 0, "count": total_managers}
            )
            
            if not all_managers:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            self.logger.info(f"Total MT5 managers: {len(all_managers)}")

            # Process each Venus group
            update_all_list = []
            
            for grp_row in groups_rows:
                grp_id = grp_row["id"]
                grp_master = grp_row["master"]
                grp_name = grp_row["name"]

                # Get login IDs for this group
                grp_logins = [
                    str(row["loginid"]) 
                    for row in group_mapping_rows 
                    if row["groupid"] == grp_id
                ]

                # Find MT5 managers matching these logins
                matching_managers = [
                    mgr for mgr in all_managers
                    if str(mgr.get("Login", "")) in grp_logins
                ]

                self.logger.debug(
                    f"Group {grp_name}: {len(matching_managers)} matching managers"
                )

                group_total_pl = 0.0

                for mgr in matching_managers:
                    mgr_login = mgr.get("Login")
                    mgr_groups = mgr.get("Groups", [])
                    mt5_group_list = [g["Group"] for g in mgr_groups]

                    if not mt5_group_list:
                        continue

                    # Get users for this manager's groups
                    users = await mt5_service.get_users_by_groups(mt5_group_list)
                    if not users:
                        continue

                    user_df = pd.DataFrame(users)
                    user_df["commission"] = user_df.get(self.config.commission_field, "")
                    user_df["commission"] = user_df["commission"].replace(r"", 0)
                    user_df["commission"] = pd.to_numeric(user_df["commission"], errors="coerce")
                    user_df["commission"] = user_df["commission"].fillna(0) * 100
                    user_df["commission"] = user_df["commission"].astype(int)
                    user_df["Login"] = user_df["Login"].astype(np.int64)

                    # Default commission from manager
                    mgr_user = user_df[user_df["Login"] == mgr_login]
                    default_comm = mgr_user["commission"].values[0] if not mgr_user.empty else 0
                    user_df.loc[user_df["commission"] == 0, "commission"] = default_comm

                    # Get trade accounts
                    accounts = await mt5_service.get_trade_accounts_by_groups(mt5_group_list)
                    if not accounts:
                        continue

                    account_df = pd.DataFrame(accounts)
                    account_df = account_df[["Login", "Credit", "Equity"]]
                    account_df["Login"] = account_df["Login"].astype(np.int64)
                    account_df["Credit"] = account_df["Credit"].astype(float)
                    account_df["Equity"] = account_df["Equity"].astype(float)

                    merged_df = pd.merge(
                        user_df[["Login", "Name", "commission"]],
                        account_df,
                        on=["Login"]
                    )

                    if merged_df.empty:
                        continue

                    merged_df["M2M"] = merged_df["Equity"] - merged_df["Credit"]
                    merged_df["Net Amount"] = merged_df["M2M"] * ((100 - merged_df["commission"]) / 100)
                    merged_df = merged_df[merged_df["Net Amount"] != 0]

                    group_total_pl += merged_df["Net Amount"].sum()

                if group_total_pl != 0:
                    update_all_list.append({
                        "MASTER": grp_master,
                        "NAME": grp_name,
                        "Profit & Loss": round(group_total_pl, 2)
                    })

            if not update_all_list:
                pdf = await pdf_service.generate_update_all_pdf(empty_df)
                return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

            result_df = pd.DataFrame(update_all_list)
            result_df = result_df.sort_values(by=["MASTER"])

            # Add summary
            total_pl = result_df["Profit & Loss"].sum()
            summary = {"MASTER": "", "NAME": "Total", "Profit & Loss": round(total_pl, 2)}
            result_df = pd.concat([result_df, pd.DataFrame([summary])], ignore_index=True)

            result_df["Profit & Loss"] = result_df["Profit & Loss"].apply(format_currency)

            pdf = await pdf_service.generate_update_all_pdf(result_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

        except Exception as e:
            self.logger.error(
                "Venus UPDATE ALL report failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            pdf = await pdf_service.generate_update_all_pdf(empty_df)
            return await self.generate_pdf_result(pdf, "UPDATE_ALL.pdf")

    async def get_top_report(
        self,
        manager_type: str,
        manager_id: str,
        commission_type: Optional[str],
        count: int = 5,
    ) -> Dict[str, Any]:
        """Generate TOP N report."""
        self.logger.info(
            f"Generating TOP {count} report",
            manager_id=manager_id,
            manager_type=manager_type,
        )

        try:
            # Get user data
            users, groups = await self.get_user_data(manager_type, manager_id)
            if not users:
                return await self._return_empty_top_report(count)

            user_df = self.extract_commission(users, manager_id)

            # Get trade accounts
            if manager_type == ManagerType.MT5:
                accounts = await mt5_service.get_trade_accounts_by_groups(groups)
            else:
                logins = user_df["Login"].tolist()
                accounts = await mt5_service.get_trade_accounts_by_logins(logins)

            if not accounts:
                return await self._return_empty_top_report(count)

            account_df = pd.DataFrame(accounts)
            # Use Credit and Equity to calculate M2M like old bot
            account_df = account_df[["Login", "Credit", "Equity"]]
            account_df["Login"] = account_df["Login"].astype(np.int64)
            account_df["Credit"] = account_df["Credit"].astype(float)
            account_df["Equity"] = account_df["Equity"].astype(float)

            # Get positions for symbol analysis
            if manager_type == ManagerType.MT5:
                positions = await mt5_service.get_positions_by_groups(groups)
            else:
                positions = await mt5_service.get_positions_by_logins(
                    user_df["Login"].tolist()
                )

            # Merge user and account data
            result_df = pd.merge(user_df, account_df, on=["Login"])

            if result_df.empty:
                return await self._return_empty_top_report(count)

            # Calculate M2M and P&L like old bot
            result_df["M2M"] = result_df["Equity"] - result_df["Credit"]
            result_df["Partner"] = result_df["M2M"] * (result_df["commission"] / 100)
            result_df["Net Amount"] = result_df["M2M"] * ((100 - result_df["commission"]) / 100)
            
            # Round and convert
            result_df["M2M"] = result_df["M2M"].round(0).astype(int)
            result_df["Partner"] = result_df["Partner"].round(0).astype(int)
            result_df["Net Amount"] = result_df["Net Amount"].round(0).astype(int)
            
            # Filter zero M2M
            result_df = result_df[result_df["M2M"] != 0]
            
            # Select P&L based on commission_type (matches old bot logic)
            if commission_type == "ours":
                result_df["Net Profit"] = result_df["Net Amount"]
            else:
                result_df["Net Profit"] = result_df["Partner"]

            # Top clients by profit/loss (matches old bot - filter positive/negative first)
            profit_clients = result_df[result_df["Net Profit"] > 0].copy()
            losing_clients = result_df[result_df["Net Profit"] < 0].copy()
            
            # Sort and take top N
            top_profit_clients = (
                profit_clients.sort_values(by=["Net Profit"], ascending=False)
                .head(count)[["Login", "Name", "Net Profit"]]
            )
            top_losing_clients = (
                losing_clients.sort_values(by=["Net Profit"], ascending=True)
                .head(count)[["Login", "Name", "Net Profit"]]
            )

            # Format (rename to P&L to match old bot output)
            top_profit_clients = top_profit_clients.rename(columns={"Net Profit": "P&L"})
            top_losing_clients = top_losing_clients.rename(columns={"Net Profit": "P&L"})
            
            top_profit_clients["P&L"] = top_profit_clients["P&L"].apply(
                lambda x: format_currency(int(x))
            )
            top_losing_clients["P&L"] = top_losing_clients["P&L"].apply(
                lambda x: format_currency(int(x))
            )

            # Top trades (from deals - current week)
            # Matches old bot logic: Monday to Friday of current week
            top_profit_trades = pd.DataFrame(
                columns=["Login", "Symbol", "Type", "Volume", "P&L"]
            )
            top_losing_trades = pd.DataFrame(
                columns=["Login", "Symbol", "Type", "Volume", "P&L"]
            )

            # Calculate time range (Monday 00:00 to Friday 23:59 of current week)
            now = datetime.now()
            start_date = now - timedelta(days=now.weekday())  # Monday
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=4)  # Friday
            end_date = end_date.replace(hour=23, minute=59, second=59)
            
            start_ts = int(start_date.timestamp())
            end_ts = int(end_date.timestamp())

            if manager_type == ManagerType.MT5:
                deals = await mt5_service.get_deals_by_groups(groups, start_ts, end_ts)
            else:
                deals = await mt5_service.get_deals_by_logins(
                    user_df["Login"].tolist(), start_ts, end_ts
                )

            if deals:
                deal_df = pd.DataFrame(deals)
                deal_df = deal_df[["Login", "Symbol", "Action", "Volume", "Profit"]]
                
                # Filter for BUY (0) and SELL (1) only
                deal_df["Action"] = deal_df["Action"].astype(int)
                deal_df = deal_df[deal_df["Action"].isin([0, 1])]
                
                # Normalize volume
                deal_df["Volume"] = deal_df["Volume"].astype(np.int64) / 10000
                
                # Convert Action to object type first to avoid FutureWarning
                deal_df["Action"] = deal_df["Action"].astype(object)
                deal_df.loc[deal_df["Action"] == 0, "Action"] = "BUY"
                deal_df.loc[deal_df["Action"] == 1, "Action"] = "SELL"
                
                deal_df = deal_df.rename(columns={"Action": "Type", "Profit": "P&L"})
                
                # Explicitly convert P&L to numeric (matches old bot logic)
                deal_df["P&L"] = deal_df["P&L"].astype(float).astype(int)

                # Filter positive/negative first, then sort and take top N (matches old bot)
                profit_trades = deal_df[deal_df["P&L"] > 0].copy()
                losing_trades = deal_df[deal_df["P&L"] < 0].copy()
                
                top_profit_trades = (
                    profit_trades.sort_values(by=["P&L"], ascending=False)
                    .head(count)[["Login", "Symbol", "Type", "Volume", "P&L"]]
                )
                top_losing_trades = (
                    losing_trades.sort_values(by=["P&L"], ascending=True)
                    .head(count)[["Login", "Symbol", "Type", "Volume", "P&L"]]
                )

                top_profit_trades["P&L"] = top_profit_trades["P&L"].apply(
                    lambda x: format_currency(int(x))
                )
                top_losing_trades["P&L"] = top_losing_trades["P&L"].apply(
                    lambda x: format_currency(int(x))
                )

            # Top symbols
            top_profit_symbols = pd.DataFrame(columns=["Symbol", "Type", "Volume", "P&L"])
            top_losing_symbols = pd.DataFrame(columns=["Symbol", "Type", "Volume", "P&L"])

            if positions:
                pos_df = pd.DataFrame(positions)
                pos_df = pos_df[["Symbol", "Action", "Volume", "Profit"]]
                
                # Clean symbol names
                pos_df["Symbol"] = pos_df["Symbol"].str.split("/").str[0]
                
                # Process volume and action
                pos_df["Volume"] = pos_df["Volume"].astype(np.int64) / 10000
                pos_df["Action"] = pos_df["Action"].astype(int)
                # Convert to object type first to avoid FutureWarning
                pos_df["Action"] = pos_df["Action"].astype(object)
                pos_df.loc[pos_df["Action"] == 0, "Action"] = "BUY"
                pos_df.loc[pos_df["Action"] == 1, "Action"] = "SELL"
                
                # Apply commission weighting to positions
                pos_user_df = pd.DataFrame(positions)
                pos_user_df = pos_user_df[["Login", "Symbol", "Action", "Volume", "Profit"]]
                # Convert types to numeric to avoid TypeError
                pos_user_df["Action"] = pos_user_df["Action"].astype(int)
                pos_user_df["Volume"] = pos_user_df["Volume"].astype(np.int64)
                pos_user_df["Profit"] = pd.to_numeric(pos_user_df["Profit"], errors="coerce").fillna(0)
                pos_user_df["Login"] = pos_user_df["Login"].astype(np.int64)
                
                # Merge with user commission data
                pos_merged = pd.merge(user_df, pos_user_df, on=["Login"])
                
                if not pos_merged.empty:
                    # Apply commission to profit
                    if commission_type == "ours":
                        pos_merged["Weighted_Profit"] = pos_merged["Profit"] * ((100 - pos_merged["commission"]) / 100)
                    else:
                        pos_merged["Weighted_Profit"] = pos_merged["Profit"] * (pos_merged["commission"] / 100)
                    
                    # Process volume and action
                    pos_merged["Volume"] = pos_merged["Volume"].astype(np.int64) / 10000
                    pos_merged["Action"] = pos_merged["Action"].astype(int)
                    pos_merged.loc[pos_merged["Action"] == 0, "Action"] = "BUY"
                    pos_merged.loc[pos_merged["Action"] == 1, "Action"] = "SELL"
                    pos_merged["Symbol"] = pos_merged["Symbol"].str.split("/").str[0]
                    
                    # Group by symbol and aggregate
                    symbol_profits = pos_merged.groupby("Symbol").agg({
                        "Weighted_Profit": "sum",
                        "Volume": "sum"
                    }).reset_index()
                    symbol_profits = symbol_profits.rename(columns={
                        "Weighted_Profit": "P&L"
                    })
                    symbol_profits["P&L"] = symbol_profits["P&L"].round(0).astype(int)
                    symbol_profits["Volume"] = symbol_profits["Volume"].round(2)
                    
                    # Determine Type based on net position
                    symbol_profits["Type"] = "BUY"
                    symbol_profits.loc[symbol_profits["Volume"] < 0, "Type"] = "SELL"
                    symbol_profits["Volume"] = symbol_profits["Volume"].abs()
                else:
                    symbol_profits = pd.DataFrame(columns=["Symbol", "Type", "Volume", "P&L"])
                
                if not symbol_profits.empty:
                    # Filter positive/negative first, then sort and take top N (matches old bot)
                    profit_symbols = symbol_profits[symbol_profits["P&L"] > 0].copy()
                    losing_symbols = symbol_profits[symbol_profits["P&L"] < 0].copy()
                    
                    top_profit_symbols = (
                        profit_symbols.sort_values(by=["P&L"], ascending=False)
                        .head(count)[["Symbol", "Type", "Volume", "P&L"]]
                    )
                    top_losing_symbols = (
                        losing_symbols.sort_values(by=["P&L"], ascending=True)
                        .head(count)[["Symbol", "Type", "Volume", "P&L"]]
                    )

                    top_profit_symbols["P&L"] = top_profit_symbols["P&L"].apply(
                        lambda x: format_currency(int(x))
                    )
                    top_losing_symbols["P&L"] = top_losing_symbols["P&L"].apply(
                        lambda x: format_currency(int(x))
                    )

            data = {
                "top_profit_client": top_profit_clients,
                "top_losing_client": top_losing_clients,
                "top_profit_trade": top_profit_trades,
                "top_losing_trade": top_losing_trades,
                "top_profit_symbol": top_profit_symbols,
                "top_losing_symbol": top_losing_symbols,
            }

            pdf = await pdf_service.generate_top_report_pdf(data, count)
            return await self.generate_pdf_result(pdf, f"TOP_{count}.pdf")

        except Exception as e:
            self.logger.error(
                f"TOP {count} report failed",
                manager_id=manager_id,
                error=str(e),
            )
            return await self._return_empty_top_report(count)

    async def _return_empty_top_report(self, count: int) -> Dict[str, Any]:
        """Return empty TOP report."""
        empty_client = pd.DataFrame(columns=["Login", "Name", "P&L"])
        empty_trade = pd.DataFrame(columns=["Login", "Symbol", "Type", "Volume", "P&L"])
        empty_symbol = pd.DataFrame(columns=["Symbol", "Type", "Volume", "P&L"])

        data = {
            "top_profit_client": empty_client,
            "top_losing_client": empty_client,
            "top_profit_trade": empty_trade,
            "top_losing_trade": empty_trade,
            "top_profit_symbol": empty_symbol,
            "top_losing_symbol": empty_symbol,
        }

        pdf = await pdf_service.generate_top_report_pdf(data, count)
        return await self.generate_pdf_result(pdf, f"TOP_{count}.pdf")

