"""Microsoft OAuth2/MSAL authentication for GTD Email Tool."""

import json
import os
from pathlib import Path
from typing import Optional

import msal


CONFIG_DIR = Path.home() / ".gtd"
CONFIG_FILE = CONFIG_DIR / "config.json"
TOKEN_CACHE_FILE = CONFIG_DIR / "token_cache.bin"

SCOPES = ["Mail.ReadWrite", "Mail.Send", "offline_access"]

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def load_config() -> dict:
    """Load configuration from ~/.gtd/config.json."""
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Config file not found at {CONFIG_FILE}. "
            "Run 'gtd setup' to configure your credentials."
        )
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)


def save_config(config: dict) -> None:
    """Save configuration to ~/.gtd/config.json."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)
    # Restrict permissions to owner only
    CONFIG_FILE.chmod(0o600)


def get_token_cache() -> msal.SerializableTokenCache:
    """Load or create a serializable token cache."""
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache


def save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """Persist the token cache to disk if it has changed."""
    if cache.has_state_changed:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())
        TOKEN_CACHE_FILE.chmod(0o600)


def build_msal_app(config: dict, cache: Optional[msal.SerializableTokenCache] = None) -> msal.PublicClientApplication:
    """Build an MSAL PublicClientApplication."""
    return msal.PublicClientApplication(
        client_id=config["client_id"],
        authority=f"https://login.microsoftonline.com/{config.get('tenant_id', 'consumers')}",
        token_cache=cache,
    )


def get_access_token() -> str:
    """
    Get a valid access token, using the cache if available,
    or triggering device flow authentication if needed.
    """
    config = load_config()
    cache = get_token_cache()
    app = build_msal_app(config, cache)

    accounts = app.get_accounts()
    token = None

    if accounts:
        # Try silent acquisition first
        token = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not token:
        raise RuntimeError(
            "No valid token found. Run 'gtd setup' to authenticate, "
            "or 'gtd auth' to re-authenticate."
        )

    save_token_cache(cache)

    if "access_token" not in token:
        error = token.get("error", "unknown_error")
        description = token.get("error_description", "No description")
        raise RuntimeError(f"Failed to obtain token: {error} - {description}")

    return token["access_token"]


def authenticate_device_flow(config: dict) -> dict:
    """
    Perform device flow authentication and return the token response.
    Prints instructions to the user.
    """
    cache = get_token_cache()
    app = build_msal_app(config, cache)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to initiate device flow: {flow.get('error_description', 'Unknown error')}")

    print("\n" + "=" * 60)
    print("MICROSOFT AUTHENTICATION REQUIRED")
    print("=" * 60)
    print(f"\n{flow['message']}")
    print("\nAfter signing in, return here and press Enter to continue...")
    input()

    token = app.acquire_token_by_device_flow(flow)

    save_token_cache(cache)

    if "access_token" not in token:
        error = token.get("error", "unknown_error")
        description = token.get("error_description", "No description")
        raise RuntimeError(f"Authentication failed: {error} - {description}")

    return token


def is_configured() -> bool:
    """Check if the application has been configured."""
    return CONFIG_FILE.exists()


def is_authenticated() -> bool:
    """Check if a valid cached token exists."""
    if not TOKEN_CACHE_FILE.exists():
        return False
    try:
        config = load_config()
        cache = get_token_cache()
        app = build_msal_app(config, cache)
        accounts = app.get_accounts()
        if not accounts:
            return False
        token = app.acquire_token_silent(SCOPES, account=accounts[0])
        return token is not None and "access_token" in token
    except Exception:
        return False
