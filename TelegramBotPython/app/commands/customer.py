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

            if not positions:
                pdf = await pdf_service.generate_customer_position_pdf(
                    empty_df, f"Position of {customer_id}"
                )
                return await self.generate_pdf_result(pdf, f"Position_{customer_id}.pdf")

            df = pd.DataFrame(positions)
            df = df[["Symbol", "TimeCreate", "Action", "Volume", "PriceOpen", "PriceCurrent", "Profit"]]

            # Process
            # Format TimeCreate (matches old bot)
            df["TimeCreate"] = pd.to_datetime(df["TimeCreate"], unit='s')
            df["TimeCreate"] = df["TimeCreate"].dt.strftime('%Y-%b-%d %H:%M:%S')

            df["Volume"] = df["Volume"].astype(np.int64) / 10000
            df["Action"] = df["Action"].astype(int)
            df.loc[df["Action"] == 0, "Action"] = "BUY"
            df.loc[df["Action"] == 1, "Action"] = "SELL"

            df["PriceOpen"] = df["PriceOpen"].round(2)
            df["PriceCurrent"] = df["PriceCurrent"].round(2)
            df["Profit"] = df["Profit"].round(2)

            df = df.rename(columns={
                "TimeCreate": "Time",
                "Action": "Type",
                "PriceOpen": "Price",
                "PriceCurrent": "Cur. Price",
                "Profit": "Profit&Loss",
            })

            # Convert Profit&Loss to string for formatting
            df["Profit&Loss"] = df["Profit&Loss"].apply(lambda x: f"{x:.2f}")

            pdf = await pdf_service.generate_customer_position_pdf(
                df, f"Position of {customer_id}"
            )
            return await self.generate_pdf_result(pdf, f"Position_{customer_id}.pdf")

        except Exception as e:
            self.logger.error(
                "Customer position failed",
                customer_id=customer_id,
                error=str(e),
            )
            pdf = await pdf_service.generate_customer_position_pdf(
                empty_df, f"Position of {customer_id}"
            )
            return await self.generate_pdf_result(pdf, f"Position_{customer_id}.pdf")

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
            columns=["Symbol", "Type", "Volume", "Price", "Profit&Loss"]
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

            deals = await mt5_service.get_deals_by_logins(
                [customer_id], from_ts, to_ts
            )

            if not deals:
                title = f"{'Today' if period == 'today' else 'This Week'}'s Trades of {customer_id}"
                pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
                return await self.generate_pdf_result(pdf, f"Trades_{customer_id}.pdf")

            df = pd.DataFrame(deals)
            
            # Filter for actual trades (not deposits/withdrawals)
            df = df[df["Action"].isin([0, 1])]

            if df.empty:
                title = f"{'Today' if period == 'today' else 'This Week'}'s Trades of {customer_id}"
                pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
                return await self.generate_pdf_result(pdf, f"Trades_{customer_id}.pdf")

            df = df[["Symbol", "Action", "Volume", "Price", "Profit"]]

            # Process
            df["Volume"] = df["Volume"].astype(np.int64) / 10000
            df["Action"] = df["Action"].astype(int)
            df.loc[df["Action"] == 0, "Action"] = "BUY"
            df.loc[df["Action"] == 1, "Action"] = "SELL"

            df["Price"] = df["Price"].round(2)
            df["Profit"] = df["Profit"].round(2)

            df = df.rename(columns={
                "Action": "Type",
                "Profit": "Profit&Loss",
            })

            df["Profit&Loss"] = df["Profit&Loss"].apply(lambda x: f"{x:.2f}")

            title = f"{'Today' if period == 'today' else 'This Week'}'s Trades of {customer_id}"
            pdf = await pdf_service.generate_customer_position_pdf(df, title)
            return await self.generate_pdf_result(pdf, f"Trades_{customer_id}.pdf")

        except Exception as e:
            self.logger.error(
                "Customer trades failed",
                customer_id=customer_id,
                period=period,
                error=str(e),
            )
            title = f"{'Today' if period == 'today' else 'This Week'}'s Trades of {customer_id}"
            pdf = await pdf_service.generate_customer_position_pdf(empty_df, title)
            return await self.generate_pdf_result(pdf, f"Trades_{customer_id}.pdf")

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

            deals = await mt5_service.get_deals_by_logins(
                [customer_id], from_ts, to_ts
            )

            if not deals:
                empty_df = pd.DataFrame(columns=["Symbol", "Volume", "Price", "P&L"])
                bills = [{
                    "symbol": "No Data",
                    "buy_df": empty_df,
                    "sell_df": empty_df,
                    "result": "N/A",
                    "grand_total": "0",
                }]
                title = f"{'Today' if period == 'today' else 'This Week'}'s Bill of {customer_id}"
                pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
                return await self.generate_pdf_result(pdf, f"Bill_{customer_id}.pdf")

            df = pd.DataFrame(deals)
            df = df[df["Action"].isin([0, 1])]

            if df.empty:
                empty_df = pd.DataFrame(columns=["Symbol", "Volume", "Price", "P&L"])
                bills = [{
                    "symbol": "No Data",
                    "buy_df": empty_df,
                    "sell_df": empty_df,
                    "result": "N/A",
                    "grand_total": "0",
                }]
                title = f"{'Today' if period == 'today' else 'This Week'}'s Bill of {customer_id}"
                pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
                return await self.generate_pdf_result(pdf, f"Bill_{customer_id}.pdf")

            # Group by symbol
            bills = []
            symbols = df["Symbol"].unique()

            for symbol in symbols:
                symbol_df = df[df["Symbol"] == symbol]

                # Separate buy and sell
                buy_df = symbol_df[symbol_df["Action"] == 0][["Time", "Volume", "Price", "Profit"]]
                sell_df = symbol_df[symbol_df["Action"] == 1][["Time", "Volume", "Price", "Profit"]]

                # Process volumes
                buy_df = buy_df.copy()
                sell_df = sell_df.copy()
                
                # Format Date (Time)
                buy_df["Time"] = pd.to_datetime(buy_df["Time"], unit='s')
                buy_df["Date"] = buy_df["Time"].dt.strftime('%d.%m %H:%M:%S')
                
                sell_df["Time"] = pd.to_datetime(sell_df["Time"], unit='s')
                sell_df["Date"] = sell_df["Time"].dt.strftime('%d.%m %H:%M:%S')

                buy_df["Volume"] = buy_df["Volume"].astype(np.int64) / 10000
                sell_df["Volume"] = sell_df["Volume"].astype(np.int64) / 10000

                buy_df = buy_df.rename(columns={"Profit": "P&L"})
                sell_df = sell_df.rename(columns={"Profit": "P&L"})
                
                # Reorder columns to match old bot: ['P&L', 'Volume', 'Price', 'Date']
                # Note: The old bot code had this order in dataframe selection, 
                # but the PDF generation might expect specific columns.
                # We will ensure these columns exist.
                buy_df = buy_df[["P&L", "Volume", "Price", "Date"]]
                sell_df = sell_df[["P&L", "Volume", "Price", "Date"]]

                # Calculate totals
                total_pl = symbol_df["Profit"].sum()
                result = "Profit" if total_pl >= 0 else "Loss"

                bills.append({
                    "symbol": symbol,
                    "buy_df": buy_df,
                    "sell_df": sell_df,
                    "result": result,
                    "grand_total": format_currency(abs(total_pl)),
                })

            title = f"{'Today' if period == 'today' else 'This Week'}'s Bill of {customer_id}"
            pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
            return await self.generate_pdf_result(pdf, f"Bill_{customer_id}.pdf")

        except Exception as e:
            self.logger.error(
                "Customer bill failed",
                customer_id=customer_id,
                period=period,
                error=str(e),
            )
            empty_df = pd.DataFrame(columns=["Symbol", "Volume", "Price", "P&L"])
            bills = [{
                "symbol": "Error",
                "buy_df": empty_df,
                "sell_df": empty_df,
                "result": "N/A",
                "grand_total": "0",
            }]
            title = f"{'Today' if period == 'today' else 'This Week'}'s Bill of {customer_id}"
            pdf = await pdf_service.generate_customer_bill_pdf(bills, title)
            return await self.generate_pdf_result(pdf, f"Bill_{customer_id}.pdf")

