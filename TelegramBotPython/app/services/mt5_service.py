"""
MT5 API Service - Async HTTP client for MT5 WebAPI.

Provides connection pooling, token caching, and retry logic.
"""

import hashlib
import secrets
from typing import Any, Dict, List, Optional

import httpx

from app.core.config import settings
from app.core.logging import get_logger
from app.core.exceptions import MT5ConnectionError, MT5AuthenticationError
from app.utils.constants import Timeouts, RetryConfig

logger = get_logger(__name__)


class MT5Service:
    """
    Async MT5 WebAPI client with connection pooling and retry logic.
    
    Implements singleton pattern for shared client across the application.
    """

    _instance: Optional["MT5Service"] = None
    _client: Optional[httpx.AsyncClient] = None
    _session_valid: bool = False

    def __new__(cls) -> "MT5Service":
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """Initialize the MT5 service."""
        self.server = settings.MT5_SERVER
        self.login = settings.MT5_LOGIN
        self.password = settings.MT5_PASSWORD
        self.base_url = f"https://{self.server}"

    async def initialize(self) -> None:
        """Initialize the HTTP client."""
        if self._client is not None:
            return

        self._client = httpx.AsyncClient(
            base_url=self.base_url,
            timeout=httpx.Timeout(Timeouts.MT5_API),
            verify=False,  # MT5 uses self-signed certificates
            headers={
                "Connection": "keep-alive",
                "User-Agent": "Mozilla/5.0",
            },
        )
        logger.info("MT5 HTTP client initialized", server=self.server)

    async def close(self) -> None:
        """Close the HTTP client."""
        if self._client is not None:
            await self._client.aclose()
            self._client = None
            self._session_valid = False
            logger.info("MT5 HTTP client closed")

    async def _ensure_authenticated(self) -> None:
        """Ensure we have a valid session."""
        if self._session_valid:
            return

        await self._authenticate()

    async def _authenticate(self) -> str:
        """
        Perform MT5 two-step authentication.
        
        Returns:
            cli_rand_answer from server
            
        Raises:
            MT5AuthenticationError: If authentication fails
        """
        if self._client is None:
            await self.initialize()

        try:
            # Step 1: Get server random
            url = f"/auth_start?version=2980&agent=WebManager&login={self.login}&type=manager"
            response = await self._client.get(url)
            response.raise_for_status()
            try:
                data = response.json()
            except Exception as e:
                logger.error("Failed to decode JSON response", endpoint="auth_start", content=response.text[:1000], error=str(e))
                raise MT5AuthenticationError(f"Invalid JSON response during auth step 1: {e}")
            srv_rand = data.get("srv_rand", "")
            
            if not srv_rand:
                raise MT5AuthenticationError("Failed to get server random")

            # Step 2: Compute answer and authenticate
            # Encode password as UTF-16LE
            encode_password = self.password.encode("utf-16-le")
            hash_password = hashlib.md5(encode_password)
            hex_hash_password = hash_password.hexdigest()
            byte_hex_hash_password = bytes.fromhex(hex_hash_password)

            # Add "WebAPI" marker
            encode_webapi = "WebAPI".encode("utf-8")
            byte_password_webapi = byte_hex_hash_password + encode_webapi

            hash_password_webapi = hashlib.md5(byte_password_webapi)
            hex_hash_password_webapi = hash_password_webapi.hexdigest()
            byte_hex_hash_password_webapi = bytes.fromhex(hex_hash_password_webapi)

            # Combine with server random
            byte_srv_rand = bytes.fromhex(srv_rand)
            byte_srv_rand_answer = byte_hex_hash_password_webapi + byte_srv_rand

            hash_byte_srv_rand_answer = hashlib.md5(byte_srv_rand_answer)
            srv_rand_answer = hash_byte_srv_rand_answer.hexdigest()
            cli_rand = secrets.token_hex(16)

            # Send authentication answer
            url = f"/auth_answer?srv_rand_answer={srv_rand_answer}&cli_rand={cli_rand}"
            response = await self._client.get(url)
            response.raise_for_status()
            try:
                data = response.json()
            except Exception as e:
                logger.error("Failed to decode JSON response", endpoint="auth_answer", content=response.text[:1000], error=str(e))
                raise MT5AuthenticationError(f"Invalid JSON response during auth step 2: {e}")
            
            cli_rand_answer = data.get("cli_rand_answer", "")
            if not cli_rand_answer:
                raise MT5AuthenticationError("Authentication failed - invalid response")

            self._session_valid = True
            logger.debug("MT5 authentication successful")
            return cli_rand_answer

        except httpx.HTTPError as e:
            logger.error("MT5 authentication HTTP error", error=str(e))
            raise MT5AuthenticationError(f"HTTP error during authentication: {e}")
        except Exception as e:
            logger.error("MT5 authentication failed", error=str(e))
            raise MT5AuthenticationError(f"Authentication failed: {e}")

    async def _request(
        self,
        endpoint: str,
        params: Optional[Dict[str, Any]] = None,
        retry_count: int = 0,
    ) -> Any:
        """
        Make an authenticated request to MT5 API.
        
        Args:
            endpoint: API endpoint (without base URL)
            params: Query parameters
            retry_count: Current retry attempt
            
        Returns:
            Response data
        """
        await self._ensure_authenticated()

        try:
            url = endpoint
            if params:
                param_str = "&".join(f"{k}={v}" for k, v in params.items())
                url = f"{endpoint}?{param_str}"

            response = await self._client.get(url)
            
            # Raise for status to catch 4xx/5xx errors
            response.raise_for_status()
            
            try:
                data = response.json()
            except Exception as e:
                logger.error(
                    "Failed to decode JSON response", 
                    endpoint=endpoint, 
                    content=response.text[:1000],  # Log first 1000 chars
                    error=str(e)
                )
                raise MT5ConnectionError(f"Invalid JSON response from MT5 API: {e}")
            
            return data.get("answer", data)

        except httpx.HTTPError as e:
            if retry_count < RetryConfig.MAX_RETRIES:
                logger.warning(
                    "MT5 request failed, retrying",
                    endpoint=endpoint,
                    retry=retry_count + 1,
                    error=str(e),
                )
                self._session_valid = False
                return await self._request(endpoint, params, retry_count + 1)
            
            raise MT5ConnectionError(f"MT5 API request failed: {e}")

    # =========================================================================
    # User APIs
    # =========================================================================

    async def get_user_detail(self, login_id: int) -> Dict[str, Any]:
        """Get user details by login ID."""
        return await self._request(f"/api/user/get", {"login": login_id})

    async def get_users_by_logins(self, login_ids: List[int]) -> List[Dict[str, Any]]:
        """Get multiple users by login IDs."""
        if not login_ids:
            return []
        logins_str = ",".join(str(x) for x in login_ids)
        return await self._request("/api/user/get_batch", {"login": logins_str})

    async def get_users_by_groups(self, groups: List[str]) -> List[Dict[str, Any]]:
        """Get all users in specified groups."""
        if not groups:
            return []
        groups_str = ",".join(groups)
        return await self._request("/api/user/get_batch", {"group": groups_str})

    async def get_user_logins_by_group(self, group: str) -> List[int]:
        """Get list of user logins in a group."""
        return await self._request("/api/user/logins", {"group": group})

    # =========================================================================
    # Manager APIs
    # =========================================================================

    async def get_manager_detail(self, login_id: int) -> List[Dict[str, Any]]:
        """Get manager details by login ID."""
        return await self._request("/api/manager/get", {"login": login_id})

    # =========================================================================
    # Trade Account APIs
    # =========================================================================

    async def get_trade_accounts_by_logins(
        self, login_ids: List[int]
    ) -> List[Dict[str, Any]]:
        """Get trade account status for multiple logins."""
        if not login_ids:
            return []
        logins_str = ",".join(str(x) for x in login_ids)
        return await self._request("/api/user/account/get_batch", {"login": logins_str})

    async def get_trade_accounts_by_groups(
        self, groups: List[str]
    ) -> List[Dict[str, Any]]:
        """Get trade account status for all users in groups."""
        if not groups:
            return []
        groups_str = ",".join(groups)
        return await self._request("/api/user/account/get_batch", {"group": groups_str})

    async def _get_all_paged(
        self,
        endpoint: str,
        params: Dict[str, Any],
        limit: int = 100
    ) -> List[Dict[str, Any]]:
        """Helper to fetch all records using pagination."""
        all_items = []
        offset = 0
        
        while True:
            current_params = params.copy()
            current_params.update({"offset": offset, "total": limit})
            
            items = await self._request(endpoint, current_params)
            if not items:
                break
                
            all_items.extend(items)
            
            if len(items) < limit:
                break
                
            offset += len(items)
            
        return all_items

    # =========================================================================
    # Position APIs
    # =========================================================================

    async def get_positions_by_logins(
        self, login_ids: List[int]
    ) -> List[Dict[str, Any]]:
        """Get open positions for multiple logins."""
        if not login_ids:
            return []
            
        # For single user, use pagination to ensure we get ALL positions
        if len(login_ids) == 1:
            return await self._get_all_paged(
                "/api/position/get_page",
                {"login": login_ids[0]}
            )
            
        logins_str = ",".join(str(x) for x in login_ids)
        return await self._request("/api/position/get_batch", {"login": logins_str})

    async def get_positions_by_groups(self, groups: List[str]) -> List[Dict[str, Any]]:
        """Get open positions for all users in groups."""
        if not groups:
            return []
        groups_str = ",".join(groups)
        return await self._request("/api/position/get_batch", {"group": groups_str})

    async def get_position_total(self, login_id: int) -> Dict[str, Any]:
        """Get total number of positions for a login."""
        return await self._request("/api/position/get_total", {"login": login_id})

    async def get_positions_paged(
        self, login_id: int, offset: int = 0, total: int = 100
    ) -> List[Dict[str, Any]]:
        """Get paginated positions for a login."""
        return await self._request(
            "/api/position/get_page",
            {"login": login_id, "offset": offset, "total": total},
        )

    # =========================================================================
    # Deal APIs
    # =========================================================================

    async def get_deals_by_logins(
        self,
        login_ids: List[int],
        from_timestamp: int,
        to_timestamp: int,
    ) -> List[Dict[str, Any]]:
        """Get deals for multiple logins within a time period."""
        if not login_ids:
            return []

        # For single user, use pagination to ensure we get ALL deals
        if len(login_ids) == 1:
            return await self._get_all_paged(
                "/api/deal/get_page",
                {
                    "login": login_ids[0],
                    "from": from_timestamp,
                    "to": to_timestamp
                }
            )

        logins_str = ",".join(str(x) for x in login_ids)
        return await self._request(
            "/api/deal/get_batch",
            {"login": logins_str, "from": from_timestamp, "to": to_timestamp},
        )

    async def get_deals_by_groups(
        self,
        groups: List[str],
        from_timestamp: int,
        to_timestamp: int,
    ) -> List[Dict[str, Any]]:
        """Get deals for all users in groups within a time period."""
        if not groups:
            return []
        groups_str = ",".join(groups)
        return await self._request(
            "/api/deal/get_batch",
            {"group": groups_str, "from": from_timestamp, "to": to_timestamp},
        )

    async def get_deal_total(
        self,
        login_id: int,
        from_timestamp: Optional[int] = None,
        to_timestamp: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Get total number of deals for a login."""
        params = {"login": login_id}
        if from_timestamp:
            params["from"] = from_timestamp
        if to_timestamp:
            params["to"] = to_timestamp
        return await self._request("/api/deal/get_total", params)

    async def get_deals_paged(
        self,
        login_id: int,
        offset: int = 0,
        total: int = 100,
        from_timestamp: Optional[int] = None,
        to_timestamp: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Get paginated deals for a login."""
        params = {"login": login_id, "offset": offset, "total": total}
        if from_timestamp:
            params["from"] = from_timestamp
        if to_timestamp:
            params["to"] = to_timestamp
        return await self._request("/api/deal/get_page", params)

    # =========================================================================
    # Symbol APIs
    # =========================================================================

    async def get_symbol(self, symbol: str) -> Dict[str, Any]:
        """Get symbol details."""
        return await self._request("/api/symbol/get", {"symbol": symbol})

    async def get_symbol_list(self) -> List[str]:
        """Get list of all symbols."""
        return await self._request("/api/symbol/list")


# Global MT5 service instance
mt5_service = MT5Service()

