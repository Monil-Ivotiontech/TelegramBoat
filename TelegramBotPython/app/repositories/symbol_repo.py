"""
Symbol Repository - Database access for symbol categories.
"""

from typing import Dict, List, Set

from app.core.database import db_pool
from app.core.logging import get_logger

logger = get_logger(__name__)


class SymbolRepository:
    """Repository for symbol category database operations."""

    _cache: Dict[str, Set[str]] = {}

    async def get_symbol_categories(self) -> Dict[str, Set[str]]:
        """
        Get all symbol categories from the database.

        Returns:
            Dictionary mapping category to set of symbols
        """
        if self._cache:
            return self._cache

        query = """
            SELECT symbol, category
            FROM symbol_category_detail
            WHERE utilflag = 'active'
        """
        
        rows = await db_pool.fetch(query)
        
        categories: Dict[str, Set[str]] = {}
        for row in rows:
            category = row["category"]
            symbol = row["symbol"]
            
            if category not in categories:
                categories[category] = set()
            categories[category].add(symbol)
        
        self._cache = categories
        return categories

    async def get_cmx_symbols(self) -> Set[str]:
        """Get all CMX symbols."""
        categories = await self.get_symbol_categories()
        return categories.get("CMX", set())

    async def get_mcx_symbols(self) -> Set[str]:
        """Get all MCX symbols."""
        categories = await self.get_symbol_categories()
        return categories.get("MCX", set())

    async def get_nse_symbols(self) -> Set[str]:
        """Get all NSE symbols."""
        categories = await self.get_symbol_categories()
        return categories.get("NSE", set())

    async def categorize_symbol(self, symbol: str) -> str:
        """
        Determine the category of a symbol.

        Args:
            symbol: Symbol name

        Returns:
            Category name ("CMX", "MCX", "NSE", or "OTHER")
        """
        categories = await self.get_symbol_categories()
        
        for category, symbols in categories.items():
            if symbol in symbols:
                return category
        
        return "OTHER"

    async def categorize_dataframe(self, df, symbol_column: str = "Symbol"):
        """
        Split a DataFrame by symbol category.

        Args:
            df: DataFrame with symbols
            symbol_column: Name of the symbol column

        Returns:
            Tuple of (cmx_df, mcx_df, other_df)
        """
        import pandas as pd
        
        if df.empty:
            empty = pd.DataFrame()
            return empty, empty, empty
        
        cmx_symbols = await self.get_cmx_symbols()
        mcx_symbols = await self.get_mcx_symbols()
        
        cmx_df = df[df[symbol_column].isin(cmx_symbols)]
        mcx_df = df[df[symbol_column].isin(mcx_symbols)]
        
        other_df = df[~df[symbol_column].isin(cmx_symbols)]
        other_df = other_df[~other_df[symbol_column].isin(mcx_symbols)]
        
        return cmx_df, mcx_df, other_df

    def clear_cache(self) -> None:
        """Clear the symbol category cache."""
        self._cache.clear()


# Global symbol repository instance
symbol_repo = SymbolRepository()

