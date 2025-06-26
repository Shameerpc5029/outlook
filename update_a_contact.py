from typing import Dict, Optional, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger  # IMPORTANT: This import is mandatory
import os


class OutlookContactUpdater:
    @staticmethod
    def update_contact_by_id(
        access_token: str, contact_id: str, update_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update an existing contact in Outlook by its ID.

        :param access_token: OAuth2 access token for Microsoft Graph API.
        :param contact_id: ID of the contact to update.
        :param update_data: Dictionary containing the fields to update.
        :return: Dictionary containing the result or error.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            logger.info(
                "Sending PATCH request to update contact.",
                extra={"url": url, "headers": headers, "update_data": update_data},
            )

            response = requests.patch(
                url, headers=headers, json=update_data, timeout=10
            )
            response.raise_for_status()

            logger.info(
                "Successfully updated contact.",
                extra={"contact_id": contact_id, "response": response.json()},
            )

            return {"result": response.json(), "error": None}
        except requests.exceptions.RequestException as e:
            logger.error(
                "API request failed.",
                extra={
                    "contact_id": contact_id,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}
        except Exception as e:
            logger.error(
                "Error updating contact.",
                extra={
                    "contact_id": contact_id,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}


def main(
    connection_id: str,
    contact_id: str,
    update_data: Dict[str, Any] ,
) -> dict[str, Any]:
    """
    Main function to update an existing contact in Outlook.

    :param connection_id: Nango connection ID for the Outlook account.
    :param contact_id: ID of the contact to update.
    :param update_data: Dictionary containing the fields to update.
    :return: Dictionary containing the result or error.
    """
    try:
        if not contact_id:
            error_message = "Contact ID is required but was not provided."
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}

        if not update_data:
            error_message = "Update data is required but was not provided."
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}

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
            "Starting contact update process.",
            extra={"contact_id": contact_id, "update_data": update_data},
        )

        contact_updater = OutlookContactUpdater()
        result = contact_updater.update_contact_by_id(
            access_token=access_token, contact_id=contact_id, update_data=update_data
        )
        return result
    except Exception as e:
        logger.error(
            "Error in main function.",
            extra={"error": str(e), "path": os.getenv("WM_JOB_PATH")},
        )
        return {"result": None, "error": str(e)}
