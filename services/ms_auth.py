import msal
import time
import logging
from typing import Optional

logger = logging.getLogger(__name__)

class TokenManager:
    def __init__(self, client_id: str, authority: str, client_secret: str, scopes: list[str]):
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            authority=authority,
            client_credential=client_secret
        )
        self._scopes = scopes
        self._token: Optional[str] = None
        self._expires_at: float = 0.0
    
    def get_token(self) -> str:
        """Returns a valid token, either from cache or by fetching a new one."""
        # 1. Check if the current token is valid (expired_on - 5 minutes > current time).
        if self._token and (self._expires_at - 300) > time.time():
            logger.info("Using cached access token.")
            return self._token
            
        logger.info("Access token is expired or missing. Acquiring new token...")
        result = self._app.acquire_token_for_client(scopes=self._scopes)

        # # 2. Try to get token from MSAL's internal cache first
        # # MSAL has its own cache! This is a "double layer" of protection.
        # accounts = self._app.get_accounts()
        # result = self._app.acquire_token_silent(self._scopes, account=accounts[0]) if accounts else None

        # if not result:
        #     # 3. If silent fails, get a fresh one from Microsoft
        #     result = self._app.acquire_token_for_client(scopes=self._scopes)

        # 4. Process result
        if result and "access_token" in result:
            self._token = result['access_token']
            
            # MSAL usually provides 'expires_in' (seconds until expiry)
            # We convert it to an absolute timestamp immediately
            self._expires_at = time.time() + result.get('expires_in', 3600)
            logger.info("Token acquired successfully.")
            return self._token
        
        # 5. Robust error handling
        error_msg = result.get('error_description') if result else "No response from MSAL"
        raise Exception(f"Failed to get token: {error_msg}")    