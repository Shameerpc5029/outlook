from typing import Optional, Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookFolderUpdater:
    @staticmethod
    def update_folder_by_id(
        access_token: str,
        folder_id: str,
        display_name: Optional[str] = None,
        parent_folder_id: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Update a folder in Microsoft Outlook using Microsoft Graph API.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # Build the payload
            payload = {}
            if display_name:
                payload["displayName"] = display_name
            if parent_folder_id:
                payload["parentFolderId"] = parent_folder_id

            if not payload:
                raise ValueError("No properties provided to update the folder.")

            response = requests.patch(url, headers=headers, json=payload, timeout=10)
            response.raise_for_status()

            updated_folder = response.json()
            logger.info(f"Successfully updated folder with ID: {folder_id}")
            return {"result": updated_folder, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error updating folder: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    folder_id: str,
    display_name: Optional[str] = "",
    parent_folder_id: Optional[str] = "",
) -> Dict[str, Any]:
    """
    Main function to update an Outlook folder.
    """
    try:
        logger.info(
            "Fetching connection credentials.", extra={"connection_id": connection_id}
        )
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )

        if not credentials or "credentials" not in credentials:
            error_message = (
                "Missing 'credentials' in the response from get_connection_credentials."
            )
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "credentials_response": credentials,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        access_token = credentials["credentials"].get("access_token")
        if not access_token:
            error_message = "Access token is missing in the credentials."
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        folder_updater = OutlookFolderUpdater()
        result = folder_updater.update_folder_by_id(
            access_token=access_token,
            folder_id=folder_id,
            display_name=display_name,
            parent_folder_id=parent_folder_id,
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
