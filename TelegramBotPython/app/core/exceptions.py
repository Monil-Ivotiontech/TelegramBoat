"""
Custom Exception Classes for VenusBot.

Provides specific exceptions for different error scenarios,
enabling better error handling and user-friendly messages.
"""

from typing import Optional


class VenusBotError(Exception):
    """Base exception for all VenusBot errors."""

    def __init__(
        self,
        message: str,
        user_message: Optional[str] = None,
        details: Optional[dict] = None,
    ):
        """
        Initialize the exception.

        Args:
            message: Technical error message for logging
            user_message: User-friendly message to display in Telegram
            details: Additional context for debugging
        """
        super().__init__(message)
        self.message = message
        self.user_message = user_message or "Something went wrong. Please try again."
        self.details = details or {}

    def __str__(self) -> str:
        return self.message


class AuthorizationError(VenusBotError):
    """Raised when user is not authorized to use the bot or command."""

    def __init__(
        self,
        message: str = "User not authorized",
        user_message: str = "You are not authorized to use this bot.",
        details: Optional[dict] = None,
    ):
        super().__init__(message, user_message, details)


class CommandNotFoundError(VenusBotError):
    """Raised when a command is not found or not available for the user."""

    def __init__(
        self,
        command: str,
        message: Optional[str] = None,
        user_message: str = "Command not found.",
        details: Optional[dict] = None,
    ):
        message = message or f"Command '{command}' not found"
        details = details or {}
        details["command"] = command
        super().__init__(message, user_message, details)


class MT5ConnectionError(VenusBotError):
    """Raised when MT5 API connection fails."""

    def __init__(
        self,
        message: str = "Failed to connect to MT5 API",
        user_message: str = "Trading service is temporarily unavailable. Please try again later.",
        details: Optional[dict] = None,
    ):
        super().__init__(message, user_message, details)


class MT5AuthenticationError(MT5ConnectionError):
    """Raised when MT5 API authentication fails."""

    def __init__(
        self,
        message: str = "MT5 authentication failed",
        user_message: str = "Trading service authentication failed. Please contact support.",
        details: Optional[dict] = None,
    ):
        super().__init__(message, user_message, details)


class DatabaseError(VenusBotError):
    """Raised when database operations fail."""

    def __init__(
        self,
        message: str = "Database operation failed",
        user_message: str = "Service temporarily unavailable. Please try again.",
        details: Optional[dict] = None,
    ):
        super().__init__(message, user_message, details)


class ManagerNotFoundError(VenusBotError):
    """Raised when a manager is not found."""

    def __init__(
        self,
        manager_id: Optional[str] = None,
        message: Optional[str] = None,
        user_message: str = "Manager not found or inactive.",
        details: Optional[dict] = None,
    ):
        message = message or f"Manager '{manager_id}' not found"
        details = details or {}
        if manager_id:
            details["manager_id"] = manager_id
        super().__init__(message, user_message, details)


class LicenseExpiredError(VenusBotError):
    """Raised when user's license has expired."""

    def __init__(
        self,
        message: str = "License expired",
        user_message: str = "Your license has expired. Please contact support to renew.",
        details: Optional[dict] = None,
    ):
        super().__init__(message, user_message, details)


class ReportGenerationError(VenusBotError):
    """Raised when report/PDF generation fails."""

    def __init__(
        self,
        report_type: str,
        message: Optional[str] = None,
        user_message: str = "Failed to generate report. Please try again.",
        details: Optional[dict] = None,
    ):
        message = message or f"Failed to generate {report_type} report"
        details = details or {}
        details["report_type"] = report_type
        super().__init__(message, user_message, details)


class InvalidInputError(VenusBotError):
    """Raised when user input is invalid."""

    def __init__(
        self,
        field: str,
        message: Optional[str] = None,
        user_message: Optional[str] = None,
        details: Optional[dict] = None,
    ):
        message = message or f"Invalid input for field '{field}'"
        user_message = user_message or f"Invalid input. Please check your {field} and try again."
        details = details or {}
        details["field"] = field
        super().__init__(message, user_message, details)

