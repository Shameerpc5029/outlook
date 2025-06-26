from typing import Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookFolderCreator:
    @staticmethod
    def create_folder(access_token: str, folder_name: str) -> Dict[str, Any]:
        try:
            url = "https://graph.microsoft.com/v1.0/me/mailFolders"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            folder_data = {"displayName": folder_name}

            response = requests.post(url, headers=headers, json=folder_data, timeout=10)
            response.raise_for_status()

            folder = response.json()
            logger.info(f"Created folder: {folder.get('id')} with name: {folder_name}")
            return {"result": folder, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error creating folder: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(connection_id: str, folder_name: str) -> Dict[str, Any]:
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        folder_creator = OutlookFolderCreator()
        result = folder_creator.create_folder(
            access_token=access_token, folder_name=folder_name
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message)
        return {"result": None, "error": error_message}
