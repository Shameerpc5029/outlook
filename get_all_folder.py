from typing import Dict, Any, List
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookFolderFetcher:
    def __init__(self, access_token: str):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0/me/mailFolders"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

    def fetch_folders(self, include_child_folders: bool = True) -> Dict[str, Any]:
        try:
            response = requests.get(
                f"{self.base_url}?$top=100", headers=self.headers, timeout=30
            )
            response.raise_for_status()

            folders = response.json().get("value", [])

            if include_child_folders:
                for folder in folders[:]:
                    child_folders = self._get_child_folders(folder["id"])
                    if child_folders:
                        folder["childFolders"] = child_folders

            logger.info(f"Successfully fetched {len(folders)} top-level folders")
            return {"result": folders, "error": None}

        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching folders: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}

    def _get_child_folders(self, parent_folder_id: str) -> List[Dict[str, Any]]:
        try:
            response = requests.get(
                f"{self.base_url}/{parent_folder_id}/childFolders",
                headers=self.headers,
                timeout=30,
            )
            response.raise_for_status()
            return response.json().get("value", [])
        except requests.exceptions.RequestException:
            logger.warning(
                f"Failed to fetch child folders for folder {parent_folder_id}",
                extra={"path": os.getenv("WM_JOB_PATH")},
            )
            return []


def main(connection_id: str, include_child_folders: bool = True) -> Dict[str, Any]:
    """
    Main function to fetch Outlook folders.

    Args:
        connection_id: The Nango connection ID
        include_child_folders: Whether to include child folders in the response

    Returns:
        Dictionary containing either the folders list or error information
    """
    try:
        # Add more detailed logging
        logger.info(
            "Starting folder fetch process",
            extra={
                "connection_id": connection_id,
                "include_child_folders": include_child_folders,
                "path": os.getenv("WM_JOB_PATH"),
            },
        )

        # Try to get credentials and log the response
        try:
            credentials = get_connection_credentials(
                id=connection_id, providerConfigKey="outlook"
            )
            logger.info(
                "Credentials fetched successfully",
                extra={
                    "connection_id": connection_id,
                    "has_credentials": bool(credentials),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
        except Exception as cred_error:
            error_message = f"Failed to fetch credentials: {str(cred_error)}"
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "error_type": type(cred_error).__name__,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        # Validate credentials structure
        if not credentials:
            error_message = "No credentials returned from get_connection_credentials"
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        if not isinstance(credentials, dict):
            error_message = f"Unexpected credentials format: {type(credentials)}"
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        # Get access token with better error handling
        try:
            access_token = credentials.get("credentials", {}).get("access_token")
        except AttributeError as e:
            error_message = f"Invalid credentials structure: {str(e)}"
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "credentials_type": type(credentials).__name__,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        if not access_token:
            error_message = "Access token is missing in the credentials"
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "credentials_keys": list(credentials.keys())
                    if isinstance(credentials, dict)
                    else "N/A",
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        folder_fetcher = OutlookFolderFetcher(access_token)
        result = folder_fetcher.fetch_folders(
            include_child_folders=include_child_folders
        )

        return result

    except Exception as e:
        error_message = f"Error in main function: {str(e)}"
        logger.error(
            error_message,
            extra={
                "connection_id": connection_id,
                "error_type": type(e).__name__,
                "path": os.getenv("WM_JOB_PATH"),
            },
        )
        return {"result": None, "error": error_message}
