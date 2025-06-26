"""Connection utilities for Outlook MCP Server"""
import os
from typing import Any
import requests
from dotenv import load_dotenv

load_dotenv(override=True)


def get_connection_credentials() -> dict[str, Any]:
    """Get credentials from Nango"""
    connection_id = os.environ.get("NANGO_CONNECTION_ID")
    integration_id = os.environ.get("NANGO_INTEGRATION_ID")
    base_url = os.environ.get("NANGO_BASE_URL")
    secret_key = os.environ.get("NANGO_SECRET_KEY")
    
    if not all([connection_id, integration_id, base_url, secret_key]):
        raise ValueError("Missing required environment variables for Nango connection")
    
    url = f"{base_url}/connection/{connection_id}"
    params = {
        "provider_config_key": integration_id,
        "refresh_token": "true",
    }
    headers = {"Authorization": f"Bearer {secret_key}"}
    
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()  # Raise exception for bad status codes
    
    return response.json()


def get_access_token() -> str:
    """Get access token from Nango credentials"""
    credentials = get_connection_credentials()
    access_token = credentials.get("credentials", {}).get("access_token")
    if not access_token:
        raise ValueError("Access token not found in credentials")
    return access_token
