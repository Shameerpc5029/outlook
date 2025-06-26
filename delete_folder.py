from typing import Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger  # IMPORTANT: This import is mandatory
import os


class OutlookFolderDeleter:
    @staticmethod
    def delete_folder_by_id(access_token: str, folder_id: str) -> Dict[str, Any]:
        try:
            url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
            }

            logger.info(
                "Sending DELETE request to delete folder.",
                extra={"url": url, "folder_id": folder_id},
            )

            response = requests.delete(url, headers=headers, timeout=10)
            response.raise_for_status()

            logger.info(
                "Successfully deleted folder.",
                extra={"folder_id": folder_id},
            )

            return {
                "result": f"Folder with ID {folder_id} deleted successfully.",
                "error": None,
            }
        except requests.exceptions.RequestException as e:
            logger.error(
                "API request failed.",
                extra={
                    "folder_id": folder_id,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}
        except Exception as e:
            logger.error(
                "Error deleting folder.",
                extra={
                    "folder_id": folder_id,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}


def main(connection_id: str, folder_id: str) -> dict[str, Any]:
    try:
        logger.info(
            "Fetching credentials for the provided connection ID.",
            extra={"connection_id": connection_id},
        )

        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )

        if "credentials" not in credentials:
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
                    "credentials": credentials,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        logger.info(
            "Starting folder deletion process.",
            extra={"folder_id": folder_id},
        )

        folder_deleter = OutlookFolderDeleter()
        result = folder_deleter.delete_folder_by_id(
            access_token=access_token, folder_id=folder_id
        )
        return result
    except Exception as e:
        logger.error(
            "Error in main function.",
            extra={"error": str(e), "path": os.getenv("WM_JOB_PATH")},
        )
        return {"result": None, "error": str(e)}
