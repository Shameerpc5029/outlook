from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookEventsFetcher:
    @staticmethod
    def get_all_events(access_token: str) -> Dict[str, Any]:
        """
        Fetch all events from the user's Outlook calendar and return only event id and subject.

        :param access_token: OAuth2 access token for Microsoft Graph API.
        :return: Dictionary containing the list of events or an error.
        """
        try:
            url = "https://graph.microsoft.com/v1.0/me/events"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            events = response.json().get("value", [])
            filtered_events = [
                {"id": event.get("id"), "name": event.get("subject")}
                for event in events
            ]

            logger.info(f"Fetched {len(filtered_events)} events.")
            return {"result": filtered_events, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching events: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(connection_id: str) -> Dict[str, Any]:
    """
    Main function to fetch all events from an Outlook calendar.

    :param connection_id: Nango connection ID for the Outlook account.
    :return: Dictionary containing the result or error.
    """
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        events_fetcher = OutlookEventsFetcher()
        result = events_fetcher.get_all_events(access_token=access_token)
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
