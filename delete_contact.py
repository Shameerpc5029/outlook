from typing import Optional, Dict, Any
import requests
from wmill import get_variable
from f.common.logfire.logger import logger  # IMPORTANT: This import is mandatory
from f.common.nango.connections import get_connection_credentials
import os


class OutlookContactDeleter:
    @staticmethod
    def delete_contact_by_id(access_token: str, contact_id: str) -> Dict[str, Any]:
        """
        Delete a contact from Outlook by its ID.

        :param access_token: OAuth2 access token for Microsoft Graph API.
        :param contact_id: ID of the contact to delete.
        :return: Dictionary containing the result or error.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.delete(url, headers=headers, timeout=10)
            if response.status_code == 204:
                logger.info(f"Successfully deleted contact with ID: {contact_id}")
                return {"result": "Contact deleted successfully", "error": None}
            else:
                error_message = f"Failed to delete contact. Status code: {response.status_code}, Response: {response.text}"
                logger.error(error_message)
                return {"result": None, "error": error_message}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error deleting contact: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(connection_id: str, contact_id: str) -> Dict[str, Any]:
    """
    Main function to delete an existing contact in Outlook.

    :param connection_id: Nango connection ID for the Outlook account.
    :param contact_id: ID of the contact to delete.
    :return: Dictionary containing the result or error.
    """
    try:
        if not contact_id:
            error_message = "Contact ID is required but was not provided."
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}

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
            "Starting contact deletion process.", extra={"contact_id": contact_id}
        )

        contact_deleter = OutlookContactDeleter()
        result = contact_deleter.delete_contact_by_id(
            access_token=access_token, contact_id=contact_id
        )
        return result
    except Exception as e:
        logger.error(
            "Error in main function.",
            extra={"error": str(e), "path": os.getenv("WM_JOB_PATH")},
        )
        return {"result": None, "error": str(e)}
