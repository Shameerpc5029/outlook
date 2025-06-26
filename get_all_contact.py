from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookContactsFetcher:
    @staticmethod
    def get_all_contacts(access_token: str) -> Dict[str, Any]:
        """
        Fetch all contacts from the user's Outlook account and return only contact id and display name.

        :param access_token: OAuth2 access token for Microsoft Graph API.
        :return: Dictionary containing the list of contacts or an error.
        """
        try:
            url = "https://graph.microsoft.com/v1.0/me/contacts"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            contacts = response.json().get("value", [])
            filtered_contacts = [
                {"id": contact.get("id"), "name": contact.get("displayName")}
                for contact in contacts
            ]

            logger.info(f"Fetched {len(filtered_contacts)} contacts.")
            return {"result": filtered_contacts, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching contacts: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(connection_id: str) -> Dict[str, Any]:
    """
    Main function to fetch all contacts from an Outlook account.

    :param connection_id: Nango connection ID for the Outlook account.
    :return: Dictionary containing the result or error.
    """
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        contacts_fetcher = OutlookContactsFetcher()
        result = contacts_fetcher.get_all_contacts(access_token=access_token)
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
