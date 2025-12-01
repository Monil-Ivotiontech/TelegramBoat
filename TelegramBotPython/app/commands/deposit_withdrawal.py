"""
Deposit/Withdrawal Command Handlers.

Handles DW report generation.
"""

from typing import Any, Dict, Optional, Tuple
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from io import BytesIO

from app.commands.base import BaseCommand
from app.core.database import db_pool
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.utils.constants import ManagerType


class DepositWithdrawalCommands(BaseCommand):
    """Handler for deposit/withdrawal reports."""

    async def get_deposit_withdrawal(
        self,
        message_text: str,
        chat_id: int,
        manager_id: str,
        manager_type: str,
    ) -> Dict[str, Any]:
        """
        Generate deposit/withdrawal report.

        OLD Format (supported):
            DW LoginID                     - Today's D/W for specific login
            DW LoginID DD/MM/YYYY          - Specific date D/W as PDF
            DW LoginID DD/MM/YYYY PDF      - Specific date D/W as PDF
            DW LoginID DD/MM/YYYY EXCEL    - Specific date D/W as Excel
        
        NEW Format (also supported):
            DW                    - Today's D/W as PDF
            DW 7                  - Last 7 days D/W as PDF
            DW EXCEL              - Today's D/W as Excel
            DW 7 EXCEL            - Last 7 days D/W as Excel
        """
        self.logger.info(
            "Generating D/W report",
            message=message_text,
            manager_id=manager_id,
        )

        try:
            # Check dwaccess permission
            has_full_access = await self._check_dw_access(chat_id, manager_id, manager_type)
            
            # Parse the command
            parsed = await self._parse_dw_command(
                message_text, 
                manager_id, 
                has_full_access
            )
            
            if parsed.get("error"):
                return self.create_text_result(parsed["error"])

            target_login = parsed["login"]
            report_date = parsed["date"]
            output_format = parsed["format"]
            commission_type = "ours" if has_full_access else "partner"

            # Calculate time range
            start_date = report_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date.replace(hour=23, minute=59, second=59)
            from_ts = int(start_date.timestamp())
            to_ts = int(end_date.timestamp())

            # Get manager details and groups
            manager_detail = await mt5_service.get_manager_detail(int(target_login))
            if not manager_detail:
                return self.create_text_result("Manager not found.")

            grp_list = manager_detail[0].get("Groups", [])
            group_list = [grp["Group"] for grp in grp_list]

            if not group_list:
                return self.create_text_result("No groups found for manager.")

            # Get users
            users = await mt5_service.get_users_by_groups(group_list)
            if not users:
                return self.create_text_result("No users found.")

            user_df = pd.DataFrame(users)
            user_df["commission"] = user_df.get(self.config.commission_field, "")
            user_df["commission"] = user_df["commission"].replace(r"", 0)
            user_df["commission"] = pd.to_numeric(user_df["commission"], errors="coerce")
            user_df["commission"] = user_df["commission"].fillna(0) * 100
            user_df["commission"] = user_df["commission"].astype(int)
            user_df["Login"] = user_df["Login"].astype(np.int64)

            # Default commission from target login
            mgr_user = user_df[user_df["Login"] == int(target_login)]
            default_comm = mgr_user["commission"].values[0] if not mgr_user.empty else 0
            user_df.loc[user_df["commission"] == 0, "commission"] = default_comm

            user_df = user_df[["Login", "Name", "Group", "commission"]]

            # Get group details for currency
            group_currency = await self._get_group_currencies()

            # Get deals
            deals = await mt5_service.get_deals_by_groups(group_list, from_ts, to_ts)
            if not deals:
                return self.create_text_result("No deposit/withdrawal found.")

            deal_df = pd.DataFrame(deals)

            # Filter for deposits and withdrawals by Comment field (matches old bot logic exactly)
            # Old bot used: deal_detail_df[(deal_detail_df["Comment"] == "Withdrawal") | (deal_detail_df["Comment"] == "Deposit")]
            deal_df = deal_df[
                (deal_df["Comment"] == "Withdrawal") | (deal_df["Comment"] == "Deposit")
            ]

            if deal_df.empty:
                return self.create_text_result("No deposit/withdrawal found.")

            deal_df = deal_df[["Login", "Profit", "Comment"]]
            deal_df["Login"] = deal_df["Login"].astype(np.int64)

            # Merge with user data
            merged_df = pd.merge(user_df, deal_df, on=["Login"])

            if merged_df.empty:
                return self.create_text_result("No deposit/withdrawal found.")

            # Calculate net amounts (matching old bot logic)
            # Use the actual Comment field from deals (already contains "Deposit" or "Withdrawal")
            merged_df["COMMENT"] = merged_df["Comment"]
            merged_df["AMOUNT USD"] = merged_df["Profit"].round(2)  # Keep sign for calculation
            
            # NET AMOUNT always uses raw commission (matches old bot)
            merged_df["NET AMOUNT"] = (
                merged_df["AMOUNT USD"] * merged_df["commission"] / 100
            ).round(2)
            
            # Now make AMOUNT USD absolute for display
            merged_df["AMOUNT USD"] = merged_df["AMOUNT USD"].abs()
            
            # RATIO for display - adjust based on commission_type
            if commission_type == "partner":
                merged_df["RATIO"] = 100 - merged_df["commission"]
            else:
                merged_df["RATIO"] = merged_df["commission"]

            # Add currency
            if not group_currency.empty:
                merged_df = pd.merge(
                    merged_df, 
                    group_currency, 
                    on="Group", 
                    how="left"
                )
                merged_df["CURRENCY"] = merged_df["Currency"].fillna("USD")
            else:
                merged_df["CURRENCY"] = "USD"

            # Add serial number and format
            merged_df = merged_df.reset_index(drop=True)
            merged_df.index = merged_df.index + 1
            merged_df["S No."] = merged_df.index

            # Format RATIO as percentage
            merged_df["RATIO"] = merged_df["RATIO"].astype(str) + "%"

            # Select and order columns
            result_df = merged_df[[
                "S No.", "Login", "Name", "COMMENT", "AMOUNT USD", 
                "RATIO", "NET AMOUNT", "CURRENCY"
            ]]
            result_df = result_df.rename(columns={"Login": "LOGIN", "Name": "NAME"})

            # Generate title
            date_str = report_date.strftime("%d/%m/%Y")
            title = f"Deposit/Withdrawal Report - {date_str}"

            if output_format == "EXCEL":
                return await self._generate_excel_result(result_df, title)
            else:
                pdf = await pdf_service.generate_deposit_withdrawal_pdf(result_df, title)
                return await self.generate_pdf_result(pdf, "DW.pdf")

        except Exception as e:
            self.logger.error(
                "D/W report failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            return self.create_text_result("Failed to generate D/W report. Please try again.")

    async def _check_dw_access(
        self,
        chat_id: int,
        manager_id: str,
        manager_type: str,
    ) -> bool:
        """Check if user has full DW access."""
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
                return False
        return True

    async def _parse_dw_command(
        self,
        message_text: str,
        default_manager_id: str,
        has_full_access: bool,
    ) -> Dict[str, Any]:
        """
        Parse DW command in both old and new formats.
        
        Returns dict with: login, date, format, error
        """
        message_text = message_text.strip()
        parts = message_text.split()
        parts = [p.strip() for p in parts if p.strip()]

        # Default values
        target_login = default_manager_id
        report_date = datetime.now()
        output_format = "PDF"

        if len(parts) == 1:
            # Just "DW" - use defaults
            return {
                "login": target_login,
                "date": report_date,
                "format": output_format,
                "error": None,
            }

        # Check if second part is a number (could be LoginID or days)
        if len(parts) >= 2:
            second_part = parts[1]
            
            # Try to parse as login ID
            try:
                potential_login = int(second_part)
                
                # Check if it's a valid login or just days
                if potential_login > 1000:  # Likely a login ID
                    if not has_full_access and str(potential_login) != default_manager_id:
                        return {"error": "Access Denied"}
                    target_login = str(potential_login)
                else:
                    # It's probably days (new format)
                    report_date = datetime.now() - timedelta(days=potential_login)
                    report_date = report_date.replace(hour=0, minute=0, second=0)
                    
            except ValueError:
                # Not a number, check for EXCEL
                if second_part.upper() == "EXCEL":
                    output_format = "EXCEL"
                elif second_part.upper() == "PDF":
                    output_format = "PDF"

        # Check for date (third part in old format)
        if len(parts) >= 3:
            third_part = parts[2]
            
            # Try to parse as date DD/MM/YYYY
            try:
                report_date = datetime.strptime(third_part, "%d/%m/%Y")
                if report_date > datetime.now():
                    return {"error": "Enter Valid Date"}
            except ValueError:
                # Not a date, check for format
                if third_part.upper() == "EXCEL":
                    output_format = "EXCEL"
                elif third_part.upper() == "PDF":
                    output_format = "PDF"

        # Check for format (fourth part)
        if len(parts) >= 4:
            fourth_part = parts[3].upper()
            if fourth_part in ("PDF", "EXCEL"):
                output_format = fourth_part

        return {
            "login": target_login,
            "date": report_date,
            "format": output_format,
            "error": None,
        }

    async def _get_group_currencies(self) -> pd.DataFrame:
        """Get group currency mappings."""
        try:
            query = """
                SELECT "Group", "Currency"
                FROM mt5_group
                WHERE utilflag = 'active'
            """
            rows = await db_pool.fetch(query)
            if rows:
                return pd.DataFrame([dict(r) for r in rows])
            return pd.DataFrame()
        except Exception:
            return pd.DataFrame()

    async def _generate_excel_result(
        self,
        df: pd.DataFrame,
        title: str,
    ) -> Dict[str, Any]:
        """Generate Excel file result."""
        try:
            output = BytesIO()
            
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name="DW", startrow=1, index=False)
                
                workbook = writer.book
                worksheet = writer.sheets["DW"]
                
                # Format header
                header_format = workbook.add_format({
                    "align": "center",
                    "valign": "vcenter",
                    "bold": True,
                    "bg_color": "#005c8f",
                    "font_color": "white",
                    "font_size": 16,
                })
                
                # Merge cells for title
                worksheet.merge_range(0, 0, 0, len(df.columns) - 1, title, header_format)
                worksheet.set_row(0, 25)
                worksheet.set_row(1, 18)
            
            output.seek(0)
            
            return {
                "type": "excel",
                "data": output,
                "filename": "DW_Report.xlsx",
            }

        except Exception as e:
            self.logger.error("Excel generation failed", error=str(e))
            # Fallback to PDF
            pdf = await pdf_service.generate_deposit_withdrawal_pdf(df, title)
            return await self.generate_pdf_result(pdf, "DW.pdf")

