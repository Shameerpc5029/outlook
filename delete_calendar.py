from typing import Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookCalendarDeleter:
    @staticmethod
    def delete_calendar_by_id(access_token: str, calendar_id: str) -> Dict[str, Any]:
        try:
            url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.delete(url, headers=headers, timeout=10)
            response.raise_for_status()

            logger.info(f"Successfully deleted calendar with ID: {calendar_id}")
            return {"result": "Calendar deleted successfully", "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error deleting calendar: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(connection_id: str, calendar_id: str) -> Dict[str, Any]:
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        calendar_deleter = OutlookCalendarDeleter()
        result = calendar_deleter.delete_calendar_by_id(
            access_token=access_token, calendar_id=calendar_id
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message)
        return {"result": None, "error": error_message}
