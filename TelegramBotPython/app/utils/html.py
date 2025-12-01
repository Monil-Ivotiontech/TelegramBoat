"""
HTML Table Generator using Jinja2 Templates.

Provides clean, maintainable HTML generation for PDF reports.
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas as pd

from app.utils.formatters import format_currency, is_negative


class HTMLGenerator:
    """
    HTML Generator for creating styled tables and reports.
    
    Uses Jinja2 templates for clean, maintainable HTML generation.
    """

    # Default styles embedded for when templates aren't available
    DEFAULT_STYLES = """
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        .report-header {
            text-align: right;
            font-weight: bold;
            margin-bottom: 10px;
        }
        .report-title {
            text-align: center;
            font-weight: bold;
            font-size: 15px;
            margin: 10px 0;
        }
        table {
            border: 1px solid black;
            border-collapse: collapse;
            margin: 10px auto;
            width: auto;
        }
        caption {
            background-color: #005c8f;
            color: white;
            text-align: center;
            font-size: 14px;
            border: 1px solid black;
            padding: 0.5rem;
            font-weight: bold;
        }
        th {
            background-color: #005c8f;
            color: white;
            text-align: center;
            font-size: 12px;
            border: 1px solid black;
            padding: 0.3rem;
        }
        td {
            border: 1px solid black;
            font-size: 11px;
            padding: 0.5rem;
            font-weight: bold;
            text-align: center;
        }
        .positive { color: green; }
        .negative { color: red; }
        .buy { color: green; }
        .sell { color: red; }
    </style>
    """

    def __init__(self, templates_dir: Optional[Path] = None):
        """
        Initialize the HTML generator.

        Args:
            templates_dir: Path to templates directory
        """
        if templates_dir and templates_dir.exists():
            self.env = Environment(
                loader=FileSystemLoader(templates_dir),
                autoescape=select_autoescape(['html', 'xml'])
            )
        else:
            self.env = None

    def generate_table(
        self,
        df: pd.DataFrame,
        caption: str,
        color_column: Optional[str] = None,
        color_logic: str = "negative",  # "negative", "buy_sell", "profit_loss"
        column_widths: Optional[Dict[str, int]] = None,
    ) -> str:
        """
        Generate an HTML table from a DataFrame.

        Args:
            df: The DataFrame to render
            caption: Table caption/title
            color_column: Column to use for row coloring
            color_logic: How to determine row colors
            column_widths: Optional column width specifications

        Returns:
            HTML string for the table
        """
        if df.empty:
            return self._empty_table(caption)

        rows_html = []
        columns = list(df.columns)

        # Generate header row
        header_cells = "".join(f'<th>{col}</th>' for col in columns)
        header = f'<thead><tr>{header_cells}</tr></thead>'

        # Generate data rows
        for _, row in df.iterrows():
            row_class = self._get_row_class(row, color_column, color_logic)
            cells = "".join(f'<td>{row[col]}</td>' for col in columns)
            rows_html.append(f'<tr class="{row_class}">{cells}</tr>')

        rows = "\n".join(rows_html)

        return f'''
        <table>
            <caption>{caption}</caption>
            {header}
            <tbody>
                {rows}
            </tbody>
        </table>
        '''

    def _get_row_class(
        self,
        row: pd.Series,
        color_column: Optional[str],
        color_logic: str,
    ) -> str:
        """Determine the CSS class for a row based on color logic."""
        if not color_column or color_column not in row.index:
            return ""

        value = row[color_column]

        if color_logic == "negative":
            return "negative" if is_negative(value) else "positive"
        elif color_logic == "buy_sell":
            return "buy" if str(value).upper() == "BUY" else "sell"
        elif color_logic == "profit_loss":
            if str(value).startswith("-"):
                return "positive"  # Negative P&L is good for broker
            return "negative"
        elif color_logic == "deposit":
            return "positive" if str(value) == "Deposit" else "negative"

        return ""

    def _empty_table(self, caption: str) -> str:
        """Generate an empty table placeholder."""
        return f'''
        <table>
            <caption>{caption}</caption>
            <tbody>
                <tr><td>No data available</td></tr>
            </tbody>
        </table>
        '''

    def wrap_document(
        self,
        content: str,
        title: str = "Report",
        include_timestamp: bool = True,
    ) -> str:
        """
        Wrap content in a complete HTML document.

        Args:
            content: HTML content to wrap
            title: Document title
            include_timestamp: Whether to include generation timestamp

        Returns:
            Complete HTML document
        """
        timestamp = ""
        if include_timestamp:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            timestamp = f'<div class="report-header">{now}</div>'

        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{title}</title>
            {self.DEFAULT_STYLES}
        </head>
        <body>
            {timestamp}
            {content}
        </body>
        </html>
        '''

    def generate_m2m_report(self, df: pd.DataFrame) -> str:
        """Generate M2M report HTML."""
        table = self.generate_table(
            df,
            caption="M2M",
            color_column="Net Amount",
            color_logic="negative",
        )
        return self.wrap_document(table, title="M2M Report")

    def generate_position_report(
        self,
        cmx_df: pd.DataFrame,
        mcx_df: pd.DataFrame,
        stock_df: pd.DataFrame,
        title: str = "COM POS",
    ) -> str:
        """Generate combined position report HTML."""
        cmx_table = self.generate_table(cmx_df, "CMX", "Type", "buy_sell")
        mcx_table = self.generate_table(mcx_df, "MCX", "Type", "buy_sell")
        stock_table = self.generate_table(stock_df, "STOCKS", "Type", "buy_sell")
        
        content = f"{cmx_table}<br>{mcx_table}<br>{stock_table}"
        return self.wrap_document(content, title=title)

    def generate_update_all_report(self, df: pd.DataFrame) -> str:
        """Generate UPDATE ALL report HTML."""
        table = self.generate_table(
            df,
            caption="UPDATE ALL",
            color_column="Profit & Loss",
            color_logic="profit_loss",
        )
        return self.wrap_document(table, title="UPDATE ALL Report")

    def generate_top_report(
        self,
        data: Dict[str, pd.DataFrame],
        count: int = 5,
    ) -> str:
        """Generate TOP N report HTML."""
        tables = []
        
        table_configs = [
            ("top_profit_client", f"TOP {count} PROFITING CLIENTS"),
            ("top_losing_client", f"TOP {count} LOSING CLIENTS"),
            ("top_profit_trade", f"TOP {count} PROFITING TRADES"),
            ("top_losing_trade", f"TOP {count} LOSING TRADES"),
            ("top_profit_symbol", f"TOP {count} PROFITING INSTRUMENTS"),
            ("top_losing_symbol", f"TOP {count} LOSING INSTRUMENTS"),
        ]
        
        for key, caption in table_configs:
            if key in data:
                tables.append(self.generate_table(data[key], caption))
        
        content = "<br><br>".join(tables)
        return self.wrap_document(content, title=f"TOP {count} Report")

    def generate_deposit_withdrawal_report(
        self,
        df: pd.DataFrame,
        title: str,
    ) -> str:
        """Generate Deposit/Withdrawal report HTML."""
        table = self.generate_table(
            df,
            caption=title,
            color_column="COMMENT",
            color_logic="deposit",
        )
        return self.wrap_document(table, title="Deposit/Withdrawal Report")

    def generate_customer_position_report(
        self,
        df: pd.DataFrame,
        title: str,
    ) -> str:
        """Generate customer position report HTML."""
        table = self.generate_table(
            df,
            caption=title.upper(),
            color_column="Profit&Loss",
            color_logic="profit_loss",
        )
        return self.wrap_document(table, title=title)

    def generate_customer_bill_report(
        self,
        bills: List[Dict[str, Any]],
        title: str,
    ) -> str:
        """Generate customer bill report HTML."""
        sections = []
        
        for bill in bills:
            buy_table = self.generate_table(bill["buy_df"], "BUY")
            sell_table = self.generate_table(bill["sell_df"], "SELL")
            
            result_class = "positive" if bill["result"] == "Profit" else "negative"
            
            section = f'''
            <div style="margin-bottom: 20px;">
                <div style="background-color: #005c8f; color: white; text-align: center;
                     font-size: 15px; padding: 0.5rem; font-weight: bold; width: 600px;
                     border: 1px solid black; margin: 0 auto;">
                    {bill["symbol"]}
                </div>
                <div style="display: flex; justify-content: center;">
                    <div>{buy_table}</div>
                    <div>{sell_table}</div>
                </div>
                <div style="display: flex; justify-content: center;">
                    <div class="{result_class}" style="text-align: center; font-size: 12px;
                         font-weight: bold; border: 1px solid black; padding: 0.5rem; width: 445px;">
                        Grand Total
                    </div>
                    <div class="{result_class}" style="text-align: center; font-size: 12px;
                         font-weight: bold; border: 1px solid black; padding: 0.5rem; width: 60px;">
                        {bill["result"]}
                    </div>
                    <div class="{result_class}" style="text-align: center; font-size: 12px;
                         font-weight: bold; border: 1px solid black; padding: 0.5rem; width: 60px;">
                        {bill["grand_total"]}
                    </div>
                </div>
            </div>
            '''
            sections.append(section)
        
        content = f'''
        <hr>
        <div class="report-title">{title.upper()}</div>
        <hr>
        {"".join(sections)}
        '''
        
        return self.wrap_document(content, title=title)


# Global HTML generator instance
html_generator = HTMLGenerator()

