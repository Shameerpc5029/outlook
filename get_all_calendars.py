from typing import Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookCalendarsFetcher:
    @staticmethod
    def get_all_calendars(access_token: str) -> Dict[str, Any]:
        """
        Fetch all calendars and return only id, name, and owner details.

        :param access_token: OAuth2 access token for Microsoft Graph API.
        :return: Dictionary containing a list of calendars or an error.
        """
        try:
            url = "https://graph.microsoft.com/v1.0/me/calendars"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            calendars = response.json().get("value", [])
            filtered_calendars = [
                {
                    "id": calendar.get("id"),
                    "name": calendar.get("name"),
                    "owner": calendar.get("owner", {}).get("name"),
                }
                for calendar in calendars
            ]

            logger.info(f"Fetched {len(filtered_calendars)} calendars.")
            return {"result": filtered_calendars, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching calendars: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(connection_id: str) -> Dict[str, Any]:
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        calendars_fetcher = OutlookCalendarsFetcher()
        result = calendars_fetcher.get_all_calendars(access_token=access_token)
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message)
        return {"result": None, "error": error_message}
