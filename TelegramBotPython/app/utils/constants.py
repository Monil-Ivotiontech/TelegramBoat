"""
Application Constants.

Centralized location for all magic numbers and constant values.
"""

from typing import Dict

# =============================================================================
# MCX Contract Sizes
# =============================================================================
# Contract sizes for MCX (Multi Commodity Exchange) symbols
MCX_CONTRACT_SIZES: Dict[str, int] = {
    "CRUDEOIL": 100,
    "GOLD": 100,
    "GOLDM": 10,
    "LEAD": 5000,
    "SILVER": 30,
    "SILVERMINI": 5,
    "ZINC": 5000,
    "NATURALGAS": 1250,
    "COPPER": 2500,
}

# =============================================================================
# Column Widths for PDF Reports
# =============================================================================
class COLUMN_WIDTHS:
    """Column width constants for PDF table generation."""
    
    LARGE = 168      # For client columns
    MEDIUM = 122     # For symbol columns  
    SMALL = 95       # For trade columns
    EXTRA_SMALL = 60 # For bill columns


# =============================================================================
# Volume Divisors
# =============================================================================
class VOLUME_DIVISORS:
    """Divisors for converting raw volume values."""
    
    VOLUME = 10_000
    VOLUME_EXT = 100_000_000


# =============================================================================
# Trade Actions
# =============================================================================
class TradeAction:
    """Trade action constants."""
    
    BUY = "BUY"
    SELL = "SELL"
    
    # Raw MT5 values
    MT5_BUY = 0
    MT5_SELL = 1


# =============================================================================
# Symbol Categories
# =============================================================================
class SymbolCategory:
    """Symbol category identifiers."""
    
    CMX = "CMX"
    MCX = "MCX"
    NSE = "NSE"


# =============================================================================
# Report Types
# =============================================================================
class ReportType:
    """Report type identifiers."""
    
    M2M = "M2M"
    COM_POS = "COM POS"
    TOTAL_POS = "TOTAL POS"
    UPDATE_ALL = "UPDATE ALL"
    TOP_5 = "TOP 5"
    TOP_10 = "TOP 10"
    DEPOSIT_WITHDRAWAL = "DW"


# =============================================================================
# Manager Types
# =============================================================================
class ManagerType:
    """Manager type identifiers."""
    
    VENUS = "venus"
    MT5 = "mt5"


# =============================================================================
# Time Periods
# =============================================================================
class TimePeriod:
    """Time period identifiers for reports."""
    
    TODAY = "today"
    WEEK = "week"


# =============================================================================
# Pagination
# =============================================================================
DEFAULT_PAGE_SIZE = 100
MAX_RESULTS = 1000


# =============================================================================
# Timeouts (in seconds)
# =============================================================================
class Timeouts:
    """Timeout values in seconds."""
    
    MT5_API = 30
    DATABASE = 20
    PDF_GENERATION = 60
    HTTP_REQUEST = 30


# =============================================================================
# Retry Configuration
# =============================================================================
class RetryConfig:
    """Retry configuration for API calls."""
    
    MAX_RETRIES = 3
    BASE_DELAY = 1  # seconds
    MAX_DELAY = 10  # seconds

