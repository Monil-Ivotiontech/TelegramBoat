"""
Formatting Utilities.

Functions for formatting numbers, currencies, and percentages.
"""

from typing import Union
from decimal import Decimal


def format_currency(amount: Union[int, float, Decimal], include_symbol: bool = False) -> str:
    """
    Format a number as currency with thousand separators.

    Args:
        amount: The amount to format
        include_symbol: Whether to include currency symbol (₹)

    Returns:
        Formatted string (e.g., "1,234,567" or "₹1,234,567")

    Examples:
        >>> format_currency(1234567)
        '1,234,567'
        >>> format_currency(-1234567.89, include_symbol=True)
        '₹-1,234,568'
    """
    try:
        # Round to integer for display
        rounded = int(round(float(amount)))
        formatted = "{:,}".format(rounded)
        
        if include_symbol:
            return f"₹{formatted}"
        return formatted
    except (ValueError, TypeError):
        return str(amount)


def format_number(
    value: Union[int, float, Decimal],
    decimal_places: int = 2,
    use_thousands_separator: bool = True,
) -> str:
    """
    Format a number with specified decimal places.

    Args:
        value: The number to format
        decimal_places: Number of decimal places (default 2)
        use_thousands_separator: Whether to use comma separators

    Returns:
        Formatted string

    Examples:
        >>> format_number(1234567.891, decimal_places=2)
        '1,234,567.89'
        >>> format_number(1234.5, decimal_places=0)
        '1,235'
    """
    try:
        value = float(value)
        
        if use_thousands_separator:
            if decimal_places == 0:
                return "{:,.0f}".format(value)
            return "{:,.{dp}f}".format(value, dp=decimal_places)
        else:
            if decimal_places == 0:
                return "{:.0f}".format(value)
            return "{:.{dp}f}".format(value, dp=decimal_places)
    except (ValueError, TypeError):
        return str(value)


def format_percentage(value: Union[int, float], include_symbol: bool = True) -> str:
    """
    Format a number as a percentage.

    Args:
        value: The percentage value (e.g., 25 for 25%)
        include_symbol: Whether to include % symbol

    Returns:
        Formatted string (e.g., "25%" or "25")

    Examples:
        >>> format_percentage(25.5)
        '25.5%'
        >>> format_percentage(100, include_symbol=False)
        '100'
    """
    try:
        value = float(value)
        
        # Check if it's a whole number
        if value == int(value):
            formatted = str(int(value))
        else:
            formatted = "{:.1f}".format(value)
        
        if include_symbol:
            return f"{formatted}%"
        return formatted
    except (ValueError, TypeError):
        return str(value)


def format_login_id(login_id: Union[int, str]) -> str:
    """
    Format a login ID for display.

    Args:
        login_id: The login ID

    Returns:
        Formatted string
    """
    return str(login_id)


def is_negative(value: Union[int, float, str]) -> bool:
    """
    Check if a value is negative.

    Args:
        value: The value to check (can be number or string)

    Returns:
        True if negative, False otherwise
    """
    if isinstance(value, str):
        return value.startswith("-")
    
    try:
        return float(value) < 0
    except (ValueError, TypeError):
        return False


def safe_float(value: Union[int, float, str, None], default: float = 0.0) -> float:
    """
    Safely convert a value to float.

    Args:
        value: The value to convert
        default: Default value if conversion fails

    Returns:
        Float value or default
    """
    if value is None:
        return default
    
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_int(value: Union[int, float, str, None], default: int = 0) -> int:
    """
    Safely convert a value to integer.

    Args:
        value: The value to convert
        default: Default value if conversion fails

    Returns:
        Integer value or default
    """
    if value is None:
        return default
    
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default

