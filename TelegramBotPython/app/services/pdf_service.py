"""
PDF Service - Async PDF generation using WeasyPrint.

Generates PDFs in memory without disk I/O.
"""

import asyncio
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO
from typing import Optional

from weasyprint import HTML, CSS

from app.core.logging import get_logger
from app.utils.html import HTMLGenerator

logger = get_logger(__name__)

# Thread pool for CPU-bound PDF generation
_pdf_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="pdf_worker")


class PDFService:
    """
    Async PDF generation service.
    
    Uses a thread pool to avoid blocking the event loop during
    CPU-intensive PDF rendering.
    """

    def __init__(self):
        """Initialize the PDF service."""
        self.html_generator = HTMLGenerator()

    async def generate_pdf(
        self,
        html_content: str,
        css: Optional[str] = None,
    ) -> BytesIO:
        """
        Generate a PDF from HTML content.

        Args:
            html_content: HTML string to convert
            css: Optional CSS string

        Returns:
            BytesIO containing the PDF data
        """
        loop = asyncio.get_event_loop()
        
        # Run PDF generation in thread pool
        pdf_bytes = await loop.run_in_executor(
            _pdf_executor,
            self._generate_pdf_sync,
            html_content,
            css,
        )
        
        return pdf_bytes

    def _generate_pdf_sync(
        self,
        html_content: str,
        css: Optional[str] = None,
    ) -> BytesIO:
        """
        Synchronous PDF generation (runs in thread pool).

        Args:
            html_content: HTML string
            css: Optional CSS string

        Returns:
            BytesIO with PDF data
        """
        try:
            pdf_buffer = BytesIO()
            
            html = HTML(string=html_content)
            
            if css:
                stylesheet = CSS(string=css)
                html.write_pdf(pdf_buffer, stylesheets=[stylesheet])
            else:
                html.write_pdf(pdf_buffer)
            
            pdf_buffer.seek(0)
            logger.debug("PDF generated successfully", size=len(pdf_buffer.getvalue()))
            
            return pdf_buffer
            
        except Exception as e:
            logger.error("PDF generation failed", error=str(e))
            raise

    async def generate_m2m_pdf(self, df) -> BytesIO:
        """Generate M2M report PDF."""
        html = self.html_generator.generate_m2m_report(df)
        return await self.generate_pdf(html)

    async def generate_position_pdf(
        self,
        cmx_df,
        mcx_df,
        stock_df,
        title: str = "COM POS",
    ) -> BytesIO:
        """Generate position report PDF."""
        html = self.html_generator.generate_position_report(
            cmx_df, mcx_df, stock_df, title
        )
        return await self.generate_pdf(html)

    async def generate_update_all_pdf(self, df) -> BytesIO:
        """Generate UPDATE ALL report PDF."""
        html = self.html_generator.generate_update_all_report(df)
        return await self.generate_pdf(html)

    async def generate_top_report_pdf(self, data: dict, count: int = 5) -> BytesIO:
        """Generate TOP N report PDF."""
        html = self.html_generator.generate_top_report(data, count)
        return await self.generate_pdf(html)

    async def generate_deposit_withdrawal_pdf(self, df, title: str) -> BytesIO:
        """Generate deposit/withdrawal report PDF."""
        html = self.html_generator.generate_deposit_withdrawal_report(df, title)
        return await self.generate_pdf(html)

    async def generate_customer_position_pdf(self, df, title: str) -> BytesIO:
        """Generate customer position report PDF."""
        html = self.html_generator.generate_customer_position_report(df, title)
        return await self.generate_pdf(html)

    async def generate_customer_bill_pdf(self, bills: list, title: str) -> BytesIO:
        """Generate customer bill report PDF."""
        html = self.html_generator.generate_customer_bill_report(bills, title)
        return await self.generate_pdf(html)

    async def generate_symbol_position_pdf(self, data: list) -> BytesIO:
        """Generate symbol position report PDF."""
        from datetime import datetime
        import pandas as pd
        
        content_parts = []
        for item in data:
            table = self.html_generator.generate_table(
                item["df"],
                caption=item["symbol"],
                color_column="Type",
                color_logic="buy_sell",
            )
            content_parts.append(table)
        
        content = "<br>".join(content_parts)
        html = self.html_generator.wrap_document(content, title="Symbol Position")
        return await self.generate_pdf(html)


# Global PDF service instance
pdf_service = PDFService()

