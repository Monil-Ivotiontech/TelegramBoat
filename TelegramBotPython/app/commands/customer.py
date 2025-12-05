"""
Customer Command Handlers.

Handles customer lookups and related reports.
"""

from typing import Any, Dict, List, Optional
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

from app.commands.base import BaseCommand
from app.core.database import db_pool
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.utils.formatters import format_currency
from app.utils.constants import TimePeriod, ManagerType


class CustomerCommands(BaseCommand):
    """Handler for customer-related commands."""

    async def try_customer_lookup(
        self,
        message_text: str,
        chat_id: int,
        manager_id: str,
        manager_type: str,
    ) -> Optional[Dict[str, Any]]:
        """
        Try to process message as customer ID lookup.

        Returns None if not a customer ID.
        """
        # Check if message is a numeric customer ID
        try:
            customer_id = int(message_text.strip())
        except ValueError:
            return None

        return await self.get_customer_info(
            customer_id, chat_id, manager_id, manager_type
        )

    async def get_customer_info(
        self,
        customer_id: int,
        chat_id: int = None,
        manager_id: str = None,
        manager_type: str = None,
    ) -> Dict[str, Any]:
        """Get customer information with action buttons."""
        self.logger.info("Looking up customer", customer_id=customer_id)

        try:
            user_detail = await mt5_service.get_user_detail(customer_id)

            if not user_detail:
                return self.create_text_result(f"Account does not Exist: {customer_id}")

            # Get account status
            accounts = await mt5_service.get_trade_accounts_by_logins([customer_id])
            
            if not accounts:
                return self.create_text_result(f"Account does not Exist: {customer_id}")
            
            account_info = accounts[0]

            # Build response
            login = user_detail.get("Login", "")
            name = user_detail.get("Name", "")
            customer_group = user_detail.get("Group", "")
            
            # Calculate P&L as Equity - Credit (matches old bot)
            equity = float(account_info.get("Equity", 0))
            credit = float(account_info.get("Credit", 0))
            profit = round(equity - credit, 2)

            # Check group access for managers with dwaccess == DISABLE (matches old bot logic)
            # This restricts which customers a manager can view based on their assigned groups
            if manager_id and manager_type == ManagerType.MT5 and chat_id:
                has_access = await self._check_customer_group_access(
                    manager_id, chat_id, customer_group
                )
                if not has_access:
                    return self.create_text_result("Access Denied")

            message = (
                f"Login : {login}\n"
                f"Name : {name}\n"
                f"P&L : {profit}\n"
                f"Credit : {credit}\n"
                f"Group : {customer_group}"
            )

            if login:
                keyboard = [
                    [InlineKeyboardButton(
                        f"Position of {login}",
                        callback_data=f"Position of {login}"
                    )],
                    [InlineKeyboardButton(
                        f"Today's Trades of {login}",
                        callback_data=f"Today's Trades of {login}"
                    )],
                    [InlineKeyboardButton(
                        f"This Week's Trades of {login}",
                        callback_data=f"This Week's Trades of {login}"
                    )],
                    [InlineKeyboardButton(
                        f"Today's Bill of {login}",
                        callback_data=f"Today's Bill of {login}"
                    )],
                    [InlineKeyboardButton(
                        f"This Week's Bill of {login}",
                        callback_data=f"This Week's Bill of {login}"
                    )],
                ]
                reply_markup = InlineKeyboardMarkup(keyboard)
                return self.create_text_result(message, reply_markup)

            return self.create_text_result(message)

        except Exception as e:
            self.logger.error(
                "Customer lookup failed",
                customer_id=customer_id,
                error=str(e),
            )
            return self.create_text_result("Customer lookup failed. Please try again.")

    async def _check_customer_group_access(
        self,
        manager_id: str,
        chat_id: int,
        customer_group: str,
    ) -> bool:
        """
        Check if manager has access to view a customer in the given group.
        
        When dwaccess == DISABLE, the manager can only view customers
        in groups that match their assigned groups (with wildcard support).
        
        Matches old bot logic from bot_command_execution_detail.py lines 1344-1358
        """
        try:
            # Check if manager has dwaccess disabled
            query = """
                SELECT cd.dwaccess
                FROM mt5_manager_chatid_detail cd
                JOIN mt5_manager_chatid c ON cd."Login" = c."Login"
                WHERE cd."Login" = $1 AND c.chatid = $2
                  AND c.utilflag = 'active' AND cd.utilflag = 'active'
            """
            row = await db_pool.fetchrow(query, int(manager_id), chat_id)
            
            if not row or row.get("dwaccess") != "DISABLE":
                # Full access - no restrictions
                return True
            
            # dwaccess is DISABLE - check group permissions
            # Get manager's allowed groups from MT5
            manager_detail = await mt5_service.get_manager_detail(int(manager_id))
            if not manager_detail:
                return False
            
            group_list = [grp["Group"] for grp in manager_detail[0].get("Groups", [])]
            
            # Check if customer's group matches any allowed group
            for grp in group_list:
                if grp == "*":
                    # Wildcard - full access
                    return True
                else:
                    # Check if customer group starts with the pattern
                    # e.g., "demo*" matches "demo_users", "demo_test", etc.
                    pattern = grp.replace("*", "")
                    if isinstance(customer_group, str) and customer_group.startswith(pattern):
                        return True
            
            # No matching group found
            self.logger.warning(
                "Customer group access denied",
                manager_id=manager_id,
                customer_group=customer_group,
                allowed_groups=group_list,
            )
            return False
            
        except Exception as e:
            self.logger.error(
                "Error checking customer group access",
                manager_id=manager_id,
                error=str(e),
            )
            # Default to deny on error
            return False

    async def get_customer_position(
        self,
        customer_id: int,
    ) -> Dict[str, Any]:
        """Get customer's open positions."""
        self.logger.info("Getting customer position", customer_id=customer_id)

        empty_df = pd.DataFrame(
            columns=["Symbol", "Time", "Type", "Volume", "Price", "Cur. Price", "Profit&Loss"]
        )

        try:
            positions = await mt5_service.get_positions_by_logins([customer_id])

            title = f"Position of {customer_id}"
            filename = f"{title.replace('/', '_')}.pdf"

            if not positions:
                pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
                return await self.generate_pdf_result(pdf, filename)

            df = pd.DataFrame(positions)
            df = df[["Symbol", "TimeCreate", "Action", "Volume", "PriceOpen", "PriceCurrent", "Profit"]]

            # Convert TimeCreate - handle both int and string timestamps
            df["TimeCreate"] = pd.to_numeric(df["TimeCreate"], errors='coerce')
            df["TimeCreate"] = pd.to_datetime(df["TimeCreate"], unit='s')
            df["TimeCreate"] = df["TimeCreate"].dt.strftime('%Y-%b-%d %H:%M:%S')

            # Convert Action to int FIRST (before any other numeric operations)
            df["Action"] = df["Action"].astype(int)

            # Convert Volume to int64 first, then divide
            df["Volume"] = df["Volume"].astype(np.int64)
            df["Volume"] = df["Volume"] / 10000

            # Replace Action values with strings (convert to object first to avoid FutureWarning)
            df["Action"] = df["Action"].astype(object)
            df.loc[df["Action"] == 0, "Action"] = "BUY"
            df.loc[df["Action"] == 1, "Action"] = "SELL"

            # Rename columns first
            df = df.rename(columns={
                "TimeCreate": "Time",
                "Action": "Type",
                "PriceOpen": "Price",
                "PriceCurrent": "Cur. Price",
                "Profit": "Profit&Loss",
            })

            # Explicit type conversions BEFORE rounding (matches old bot)
            df["Volume"] = df["Volume"].astype(float).round(2)
            df["Price"] = df["Price"].astype(float).round(2)
            df["Cur. Price"] = df["Cur. Price"].astype(float).round(2)
            df["Profit&Loss"] = df["Profit&Loss"].astype(float).astype(int)

            # Add Total row (matches old bot)
            sum_value = df["Profit&Loss"].sum()
            sum_row = pd.DataFrame([{
                "Symbol": "",
                "Time": "",
                "Type": "Total",
                "Volume": "",
                "Price": "",
                "Cur. Price": "",
                "Profit&Loss": sum_value,
            }])
            df = pd.concat([df, sum_row], ignore_index=True)

            # Format Profit&Loss with currency formatting (matches old bot)
            df["Profit&Loss"] = df["Profit&Loss"].apply(
                lambda x: format_currency(x) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)
            )

            pdf = await pdf_service.generate_customer_position_pdf(df, title)
            return await self.generate_pdf_result(pdf, filename)

        except Exception as e:
            self.logger.error(
                "Customer position failed",
                customer_id=customer_id,
                error=str(e),
            )
            pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
            return await self.generate_pdf_result(pdf, filename)

    async def get_customer_trades(
        self,
        customer_id: int,
        period: str,
    ) -> Dict[str, Any]:
        """Get customer's trades for a period."""
        self.logger.info(
            "Getting customer trades",
            customer_id=customer_id,
            period=period,
        )

        empty_df = pd.DataFrame(
            columns=["Time", "Symbol", "Type", "Entry", "Volume", "Price", "Reason", "Commission", "Profit&Loss"]
        )

        try:
            # Calculate time range (matches old bot logic)
            now = datetime.now()
            if period == TimePeriod.TODAY:
                from_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
                to_date = from_date.replace(hour=23, minute=59, second=59)
            else:  # week - This Week (Monday to Friday) like old bot
                # Start from Monday of current week
                from_date = now - timedelta(days=now.weekday() % 7)
                from_date = from_date.replace(hour=0, minute=0, second=0, microsecond=0)
                # End on Friday of current week
                to_date = from_date + timedelta(days=4)
                to_date = to_date.replace(hour=23, minute=59, second=59)

            from_ts = int(from_date.timestamp())
            to_ts = int(to_date.timestamp())

            # Set title and filename to match selected option (like old bot)
            title = f"{'Today' if period == 'today' else 'This Week'}'s Trades of {customer_id}"
            filename = f"{title.replace('/', '_')}.pdf"

            deals = await mt5_service.get_deals_by_logins(
                [customer_id], from_ts, to_ts
            )

            if not deals:
                pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
                return await self.generate_pdf_result(pdf, filename)

            df = pd.DataFrame(deals)

            # Select required columns (matches old bot)
            df = df[["Time", "Symbol", "Action", "Entry", "Volume", "Price", "Reason", "Commission", "Profit"]]

            # Convert Time - handle both int and string timestamps
            df["Time"] = pd.to_numeric(df["Time"], errors='coerce')
            df["Time"] = pd.to_datetime(df["Time"], unit='s')
            df = df.sort_values(by=['Time'], ascending=False)
            df["Time"] = df["Time"].dt.strftime('%Y-%b-%d %H:%M:%S')

            # Convert Volume to int64 first, then divide
            df["Volume"] = df["Volume"].astype(np.int64)
            df["Volume"] = df["Volume"] / 10000

            # Convert Action, Entry, Reason to int
            df["Action"] = df["Action"].astype(int)
            df["Entry"] = df["Entry"].astype(int)
            df["Reason"] = df["Reason"].astype(int)

            # Replace Action values (convert to object first to avoid FutureWarning)
            df["Action"] = df["Action"].astype(object)
            df.loc[df["Action"] == 0, "Action"] = "BUY"
            df.loc[df["Action"] == 1, "Action"] = "SELL"

            # Replace Entry values (convert to object first to avoid FutureWarning)
            df["Entry"] = df["Entry"].astype(object)
            df.loc[df["Entry"] == 0, "Entry"] = "IN"
            df.loc[df["Entry"] == 1, "Entry"] = "OUT"
            df.loc[df["Entry"] == 2, "Entry"] = "INOUT"

            # Replace Reason values (convert to object first to avoid FutureWarning)
            df["Reason"] = df["Reason"].astype(object)
            df.loc[df["Reason"] == 0, "Reason"] = "CLIENT"
            df.loc[df["Reason"] == 1, "Reason"] = "EXPERT"
            df.loc[df["Reason"] == 2, "Reason"] = "DEALER"
            df.loc[df["Reason"] == 16, "Reason"] = "MOBILE"

            # Rename columns
            df = df.rename(columns={
                "Action": "Type",
                "Profit": "Profit&Loss",
            })

            # Explicit type conversions BEFORE rounding (matches old bot)
            df["Volume"] = df["Volume"].astype(float).round(2)
            df["Price"] = df["Price"].astype(float).round(2)
            df["Commission"] = df["Commission"].astype(float).astype(int)
            df["Profit&Loss"] = df["Profit&Loss"].astype(float).astype(int)

            # Add Total row (matches old bot)
            commission_sum = df["Commission"].sum()
            profit_sum = df["Profit&Loss"].sum()
            sum_row = pd.DataFrame([{
                "Time": "",
                "Symbol": "",
                "Type": "",
                "Entry": "Total",
                "Volume": "",
                "Price": "",
                "Reason": "",
                "Commission": commission_sum,
                "Profit&Loss": profit_sum,
            }])
            df = pd.concat([df, sum_row], ignore_index=True)

            # Format Commission and Profit&Loss with currency formatting (matches old bot)
            df["Commission"] = df["Commission"].apply(
                lambda x: format_currency(x) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)
            )
            df["Profit&Loss"] = df["Profit&Loss"].apply(
                lambda x: format_currency(x) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)
            )

            pdf = await pdf_service.generate_customer_position_pdf(df, title)
            return await self.generate_pdf_result(pdf, filename)

        except Exception as e:
            self.logger.error(
                "Customer trades failed",
                customer_id=customer_id,
                period=period,
                error=str(e),
            )
            pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
            return await self.generate_pdf_result(pdf, filename)

    async def get_customer_bill(
        self,
        customer_id: int,
        period: str,
    ) -> Dict[str, Any]:
        """Get customer's bill for a period."""
        self.logger.info(
            "Getting customer bill",
            customer_id=customer_id,
            period=period,
        )

        empty_df = pd.DataFrame(columns=["P&L", "Volume", "Price", "Date"])

        try:
            # Calculate time range (matches old bot logic)
            now = datetime.now()
            if period == TimePeriod.TODAY:
                from_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
                to_date = from_date.replace(hour=23, minute=59, second=59)
            else:  # week - This Week (Monday to Friday) like old bot
                # Start from Monday of current week
                from_date = now - timedelta(days=now.weekday() % 7)
                from_date = from_date.replace(hour=0, minute=0, second=0, microsecond=0)
                # End on Friday of current week
                to_date = from_date + timedelta(days=4)
                to_date = to_date.replace(hour=23, minute=59, second=59)

            from_ts = int(from_date.timestamp())
            to_ts = int(to_date.timestamp())

            # Set title and filename to match selected option (like old bot)
            title = f"{'Today' if period == 'today' else 'This Week'}'s Bill of {customer_id}"
            filename = f"{title.replace('/', '_')}.pdf"

            deals = await mt5_service.get_deals_by_logins(
                [customer_id], from_ts, to_ts
            )

            if not deals:
                bills = [{
                    "symbol": "No Data",
                    "buy_df": empty_df,
                    "sell_df": empty_df,
                    "result": "Profit",
                    "grand_total": 0,
                }]
                pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
                return await self.generate_pdf_result(pdf, filename)

            df = pd.DataFrame(deals)
            
            # Select required columns (matches old bot)
            df = df[["Time", "Symbol", "Action", "Volume", "Price", "Profit"]]

            # Convert Time - handle both int and string timestamps
            df["Time"] = pd.to_numeric(df["Time"], errors='coerce')
            df["Time"] = pd.to_datetime(df["Time"], unit='s')
            df = df.sort_values(by=['Time'], ascending=False)
            df["Time"] = df["Time"].dt.strftime('%d.%m %H:%M:%S')

            # Convert Volume to int64 first, then divide
            df["Volume"] = df["Volume"].astype(np.int64)
            df["Volume"] = df["Volume"] / 10000

            # Convert Action to int
            df["Action"] = df["Action"].astype(int)

            # Replace Action values (convert to object first to avoid FutureWarning)
            df["Action"] = df["Action"].astype(object)
            df.loc[df["Action"] == 0, "Action"] = "BUY"
            df.loc[df["Action"] == 1, "Action"] = "SELL"

            # Rename columns
            df = df.rename(columns={"Action": "Type", "Profit": "P&L", "Time": "Date"})

            # Explicit type conversions (matches old bot)
            df["Volume"] = df["Volume"].astype(float).round(2)
            df["Price"] = df["Price"].astype(float).round(2)
            df["P&L"] = df["P&L"].astype(float).astype(int)

            # Reorder columns
            df = df[["Symbol", "Type", "P&L", "Volume", "Price", "Date"]]

            # Group by symbol (matches old bot)
            bills = []
            symbol_list = sorted(df["Symbol"].unique().tolist())

            for sym in symbol_list:
                tmp_sym_df = df[df["Symbol"] == sym]
                buy_df = tmp_sym_df[tmp_sym_df["Type"] == "BUY"][["P&L", "Volume", "Price", "Date"]].copy()
                sell_df = tmp_sym_df[tmp_sym_df["Type"] == "SELL"][["P&L", "Volume", "Price", "Date"]].copy()

                # Calculate totals
                total_buy = buy_df["P&L"].sum() if not buy_df.empty else 0
                total_sell = sell_df["P&L"].sum() if not sell_df.empty else 0
                grand_total = total_buy + total_sell

                result = "Profit" if grand_total >= 0 else "Loss"

                # Equalize DataFrame lengths (matches old bot)
                if buy_df.shape[0] > sell_df.shape[0]:
                    sell_df = sell_df.reindex(list(range(0, buy_df.shape[0]))).reset_index(drop=True)
                elif buy_df.shape[0] < sell_df.shape[0]:
                    buy_df = buy_df.reindex(list(range(0, sell_df.shape[0]))).reset_index(drop=True)

                # Add Total row to buy_df (matches old bot)
                buy_sum = buy_df["P&L"].sum() if not buy_df.empty else 0
                buy_volume_sum = buy_df["Volume"].sum() if not buy_df.empty else 0
                buy_sum_row = pd.DataFrame([{
                    "P&L": buy_sum,
                    "Volume": buy_volume_sum,
                    "Price": "",
                    "Date": "Total",
                }])
                buy_df = pd.concat([buy_df, buy_sum_row], ignore_index=True)

                # Add Total row to sell_df (matches old bot)
                sell_sum = sell_df["P&L"].sum() if not sell_df.empty else 0
                sell_volume_sum = sell_df["Volume"].sum() if not sell_df.empty else 0
                sell_sum_row = pd.DataFrame([{
                    "P&L": sell_sum,
                    "Volume": sell_volume_sum,
                    "Price": "",
                    "Date": "Total",
                }])
                sell_df = pd.concat([sell_df, sell_sum_row], ignore_index=True)

                # Format P&L with currency formatting (matches old bot)
                buy_df["P&L"] = buy_df["P&L"].apply(
                    lambda x: format_currency(x) if isinstance(x, (int, float)) and not pd.isna(x) else str(x) if pd.notna(x) else " "
                )
                sell_df["P&L"] = sell_df["P&L"].apply(
                    lambda x: format_currency(x) if isinstance(x, (int, float)) and not pd.isna(x) else str(x) if pd.notna(x) else " "
                )

                # Replace nan values with space (matches old bot)
                buy_df = buy_df.fillna(" ")
                sell_df = sell_df.fillna(" ")
                buy_df = buy_df.replace("nan", " ")
                sell_df = sell_df.replace("nan", " ")

                bills.append({
                    "symbol": sym,
                    "buy_df": buy_df,
                    "sell_df": sell_df,
                    "result": result,
                    "grand_total": grand_total,
                })

            pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
            return await self.generate_pdf_result(pdf, filename)

        except Exception as e:
            self.logger.error(
                "Customer bill failed",
                customer_id=customer_id,
                period=period,
                error=str(e),
            )
            bills = [{
                "symbol": "Error",
                "buy_df": empty_df,
                "sell_df": empty_df,
                "result": "N/A",
                "grand_total": 0,
            }]
            pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
            return await self.generate_pdf_result(pdf, filename)

