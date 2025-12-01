"""
Position Command Handlers.

Generates COM POS and TOTAL POS reports.
"""

from typing import Any, Dict, Optional, Tuple
import pandas as pd
import numpy as np

from app.commands.base import BaseCommand
from app.services.mt5_service import mt5_service
from app.services.pdf_service import pdf_service
from app.repositories.symbol_repo import symbol_repo
from app.utils.formatters import format_currency
from app.utils.constants import ManagerType, MCX_CONTRACT_SIZES


class PositionCommands(BaseCommand):
    """Handler for position reports (COM POS, TOTAL POS)."""

    async def get_com_pos_report(
        self,
        manager_type: str,
        manager_id: str,
        commission_type: Optional[str],
    ) -> Dict[str, Any]:
        """
        Generate COM POS (Commission Position) report.

        Shows positions weighted by commission percentage.
        """
        self.logger.info(
            "Generating COM POS report",
            manager_id=manager_id,
            manager_type=manager_type,
        )

        empty_df = pd.DataFrame(
            columns=["Symbol", "Type", "Volume", "Avg. Price", 
                     "Cur. Price", "Profit&Loss", "Holding Volume"]
        )

        try:
            # Get user data
            users, groups = await self.get_user_data(manager_type, manager_id)
            if not users:
                return await self._return_empty_position_report()

            # Extract commission data
            user_df = self.extract_commission(users, manager_id)

            # Get positions
            if manager_type == ManagerType.MT5:
                positions = await mt5_service.get_positions_by_groups(groups)
            else:
                logins = user_df["Login"].tolist()
                positions = await mt5_service.get_positions_by_logins(logins)

            if not positions:
                return await self._return_empty_position_report()
            # Create position DataFrame
            position_df = pd.DataFrame(positions)
            position_df = position_df[[
                "Login", "Symbol", "Action", "PriceOpen", 
                "PriceCurrent", "Volume", "VolumeExt"
            ]]
            position_df["Login"] = position_df["Login"].astype(np.int64)

            # Merge with user data
            merged_df = pd.merge(user_df, position_df, on=["Login"])

            if merged_df.empty:
                return await self._return_empty_position_report()

            # Process the data
            result_df = await self._process_position_data(
                merged_df, commission_type
            )

            # Categorize and format
            cmx_df, mcx_df, stock_df = await self._categorize_positions(result_df)

            pdf = await pdf_service.generate_position_pdf(
                cmx_df, mcx_df, stock_df, "COM POS"
            )
            return await self.generate_pdf_result(pdf, "COM_POS.pdf")

        except Exception as e:
            self.logger.error(
                "COM POS report failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            return await self._return_empty_position_report()

    async def get_total_pos_report(
        self,
        manager_type: str,
        manager_id: str,
    ) -> Dict[str, Any]:
        """
        Generate TOTAL POS report.

        Shows total positions without commission weighting.
        """
        self.logger.info(
            "Generating TOTAL POS report",
            manager_id=manager_id,
            manager_type=manager_type,
        )

        try:
            # Get positions directly
            if manager_type == ManagerType.MT5:
                from app.services.manager_service import manager_service
                groups = await manager_service.get_manager_groups(manager_id, manager_type)
                if not groups:
                    return await self._return_empty_position_report()
                positions = await mt5_service.get_positions_by_groups(groups)
            else:
                from app.services.manager_service import manager_service
                logins = await manager_service.get_manager_user_logins(manager_id, manager_type)
                if not logins:
                    return await self._return_empty_position_report()
                positions = await mt5_service.get_positions_by_logins(logins)

            if not positions:
                return await self._return_empty_position_report()

            # Create DataFrame
            position_df = pd.DataFrame(positions)
            position_df = position_df[[
                "Login", "Symbol", "Action", "PriceOpen",
                "PriceCurrent", "Volume", "VolumeExt"
            ]]

            # Process without commission weighting
            result_df = await self._process_total_position_data(position_df)

            # Categorize and format
            cmx_df, mcx_df, stock_df = await self._categorize_positions(result_df)

            pdf = await pdf_service.generate_position_pdf(
                cmx_df, mcx_df, stock_df, "TOTAL POS"
            )
            return await self.generate_pdf_result(pdf, "TOTAL_POS.pdf")

        except Exception as e:
            self.logger.error(
                "TOTAL POS report failed",
                manager_id=manager_id,
                error=str(e),
                exc_info=True,
            )
            return await self._return_empty_position_report()

    async def _return_empty_position_report(self) -> Dict[str, Any]:
        """Return empty position report PDF."""
        empty_df = pd.DataFrame(
            columns=["Symbol", "Type", "Volume", "Avg. Price",
                     "Cur. Price", "Profit&Loss", "Holding Volume"]
        )
        pdf = await pdf_service.generate_position_pdf(
            empty_df, empty_df, empty_df, "Position Report"
        )
        return await self.generate_pdf_result(pdf, "Position.pdf")

    async def _process_position_data(
        self,
        df: pd.DataFrame,
        commission_type: Optional[str],
    ) -> pd.DataFrame:
        """Process position data with commission weighting."""
        # Process volume and action
        df = self.process_volume(df)
        df = self.process_action(df)

        # Explicitly convert types to avoid TypeError
        df["PriceOpen"] = pd.to_numeric(df["PriceOpen"], errors="coerce").fillna(0)
        df["PriceCurrent"] = pd.to_numeric(df["PriceCurrent"], errors="coerce").fillna(0)
        df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(0)
        if "VolumeExt" in df.columns:
            df["VolumeExt"] = pd.to_numeric(df["VolumeExt"], errors="coerce").fillna(0)
        if "VolumeExt" in df.columns:
            df["VolumeExt"] = pd.to_numeric(df["VolumeExt"], errors="coerce").fillna(0)

        # Apply commission
        if commission_type == "ours":
            df["com_commission"] = 100 - df["commission"]
        else:
            df["com_commission"] = df["commission"]

        # Calculate commission-weighted volume
        df["com_volume"] = (df["Volume"] * df["com_commission"]) / 100
        df.loc[df["Action"] == "SELL", "com_volume"] *= -1

        # Calculate weighted price
        df["com_price"] = df["PriceOpen"] * df["com_volume"]

        # Clean symbol name
        df["Symbol"] = df["Symbol"].str.split("/").str[0]

        # Store unique symbols for current price lookup
        unique_df = df.drop_duplicates(
            subset=["Symbol", "Action"], keep="first"
        ).reset_index()

        # Aggregate by symbol
        df = df.groupby("Symbol", as_index=False).agg({
            "com_volume": "sum",
            "com_price": "sum",
        })

        # Determine action based on net volume
        df["Action"] = "BUY"
        df.loc[df["com_volume"] <= 0, "Action"] = "SELL"

        # Merge with current prices
        df = pd.merge(
            df,
            unique_df[["Symbol", "Action", "PriceCurrent"]],
            on=["Symbol", "Action"],
            how="left",
        )

        # Calculate average price and P&L
        df["Avg. Price"] = (df["com_price"] / df["com_volume"]).abs().round(2)
        df["PriceCurrent"] = df["PriceCurrent"].round(2)

        # Calculate P&L
        df.loc[df["com_volume"] < 0, "Profit&Loss"] = (
            (df["PriceCurrent"] - df["Avg. Price"]) * df["com_volume"].abs()
        )
        df.loc[df["com_volume"] >= 0, "Profit&Loss"] = (
            (df["Avg. Price"] - df["PriceCurrent"]) * df["com_volume"]
        )

        # Calculate holding volume
        df["Holding Volume"] = df["Avg. Price"] * df["com_volume"].abs()

        # Clean up
        df.replace([np.inf, -np.inf], 0, inplace=True)
        df.fillna(0, inplace=True)

        # Round and convert types
        df["Profit&Loss"] = df["Profit&Loss"].round(0).astype(int)
        df["Holding Volume"] = df["Holding Volume"].round(0).astype(int)
        df["com_volume"] = df["com_volume"].round(0).astype(int)

        # Rename columns
        df = df.rename(columns={
            "com_volume": "Volume",
            "Action": "Type",
            "PriceCurrent": "Cur. Price",
        })

        # Filter zero volume
        df = df[df["Volume"] != 0]

        return df[[
            "Symbol", "Type", "Volume", "Avg. Price",
            "Cur. Price", "Profit&Loss", "Holding Volume"
        ]]

    async def _process_total_position_data(
        self,
        df: pd.DataFrame,
    ) -> pd.DataFrame:
        """Process position data without commission weighting."""
        # Process volume and action
        df = self.process_volume(df)
        df = self.process_action(df)

        # Explicitly convert types to avoid TypeError
        df["PriceOpen"] = pd.to_numeric(df["PriceOpen"], errors="coerce").fillna(0)
        df["PriceCurrent"] = pd.to_numeric(df["PriceCurrent"], errors="coerce").fillna(0)
        df["Volume"] = pd.to_numeric(df["Volume"], errors="coerce").fillna(0)
        if "VolumeExt" in df.columns:
            df["VolumeExt"] = pd.to_numeric(df["VolumeExt"], errors="coerce").fillna(0)

        # Use full volume
        df["net_volume"] = df["Volume"]
        df.loc[df["Action"] == "SELL", "net_volume"] *= -1

        # Calculate weighted price
        df["total_price"] = df["PriceOpen"] * df["net_volume"]

        # Clean symbol name
        df["Symbol"] = df["Symbol"].str.split("/").str[0]

        # Store unique symbols
        unique_df = df.drop_duplicates(
            subset=["Symbol", "Action"], keep="first"
        ).reset_index()

        # Aggregate by symbol
        df = df.groupby("Symbol", as_index=False).agg({
            "net_volume": "sum",
            "total_price": "sum",
        })

        # Determine action
        df["Action"] = "BUY"
        df.loc[df["net_volume"] <= 0, "Action"] = "SELL"

        # Merge with current prices
        df = pd.merge(
            df,
            unique_df[["Symbol", "Action", "PriceCurrent"]],
            on=["Symbol", "Action"],
            how="left",
        )

        # Calculate values
        df["Avg. Price"] = (df["total_price"] / df["net_volume"]).abs().round(2)
        df["PriceCurrent"] = df["PriceCurrent"].round(2)

        df.loc[df["net_volume"] < 0, "Profit&Loss"] = (
            (df["PriceCurrent"] - df["Avg. Price"]) * df["net_volume"].abs()
        )
        df.loc[df["net_volume"] >= 0, "Profit&Loss"] = (
            (df["Avg. Price"] - df["PriceCurrent"]) * df["net_volume"]
        )

        df["Holding Volume"] = df["Avg. Price"] * df["net_volume"].abs()

        # Clean up
        df.replace([np.inf, -np.inf], 0, inplace=True)
        df.fillna(0, inplace=True)

        df["Profit&Loss"] = df["Profit&Loss"].round(0).astype(int)
        df["Holding Volume"] = df["Holding Volume"].round(0).astype(int)
        df["net_volume"] = df["net_volume"].round(0).astype(int)

        df = df.rename(columns={
            "net_volume": "Volume",
            "Action": "Type",
            "PriceCurrent": "Cur. Price",
        })

        df = df[df["Volume"] != 0]

        return df[[
            "Symbol", "Type", "Volume", "Avg. Price",
            "Cur. Price", "Profit&Loss", "Holding Volume"
        ]]

    async def _categorize_positions(
        self,
        df: pd.DataFrame,
    ) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Categorize positions into CMX, MCX, and other."""
        if df.empty:
            empty = pd.DataFrame(
                columns=["Symbol", "Type", "Volume", "Avg. Price",
                         "Cur. Price", "Profit&Loss", "Holding Volume"]
            )
            return empty.copy(), empty.copy(), empty.copy()

        cmx_df, mcx_df, stock_df = await symbol_repo.categorize_dataframe(df, "Symbol")

        # Format CMX
        if not cmx_df.empty:
            cmx_df = self._add_summary_and_format(cmx_df)

        # Format MCX with contract sizes
        if not mcx_df.empty:
            mcx_df = self._apply_mcx_contract_sizes(mcx_df)
            mcx_df = self._add_summary_and_format(mcx_df)

        # Format other
        if not stock_df.empty:
            stock_df = self._add_summary_and_format(stock_df)

        return cmx_df, mcx_df, stock_df

    def _apply_mcx_contract_sizes(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply MCX contract sizes to P&L and Holding Volume."""
        for idx, row in df.iterrows():
            symbol = row["Symbol"]
            if symbol in MCX_CONTRACT_SIZES:
                contract_size = MCX_CONTRACT_SIZES[symbol]

                # Recalculate P&L
                if row["Type"] == "SELL":
                    pl = (row["Avg. Price"] - row["Cur. Price"]) * abs(row["Volume"]) * contract_size
                else:
                    pl = (row["Cur. Price"] - row["Avg. Price"]) * abs(row["Volume"]) * contract_size

                df.at[idx, "Profit&Loss"] = int(round(pl))
                df.at[idx, "Holding Volume"] = int(round(
                    abs(row["Volume"]) * row["Avg. Price"] * contract_size
                ))

        return df

    def _add_summary_and_format(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add summary row and format currency columns."""
        # Add summary row
        summary = {
            "Symbol": "Total",
            "Type": "",
            "Volume": df["Volume"].sum(),
            "Avg. Price": "",
            "Cur. Price": "",
            "Profit&Loss": df["Profit&Loss"].sum(),
            "Holding Volume": df["Holding Volume"].sum(),
        }
        df = pd.concat([df, pd.DataFrame([summary])], ignore_index=True)

        # Fix type based on volume sign
        df["Type"] = "BUY"
        df.loc[df["Volume"] > 0, "Type"] = "SELL"
        df["Volume"] = df["Volume"].abs()

        # Format currency
        df["Profit&Loss"] = df["Profit&Loss"].apply(format_currency)
        df["Holding Volume"] = df["Holding Volume"].apply(format_currency)

        return df

