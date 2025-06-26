from typing import Optional, Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookEventFetcher:
    @staticmethod
    def get_events(
        access_token: str,
        calendar_id: str = "",
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Retrieve events from an Outlook calendar using Microsoft Graph API.
        Filters events by calendar ID, start, and end datetime if provided.
        """
        try:
            url = "https://graph.microsoft.com/v1.0/me/events"

            # Construct filter query if start_datetime or end_datetime are provided
            query_params = {}
            if calendar_id:
                url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events"

            if start_datetime and end_datetime:
                query_params["$filter"] = (
                    f"start/dateTime ge '{start_datetime}' and end/dateTime le '{end_datetime}'"
                )
            elif start_datetime:
                query_params["$filter"] = f"start/dateTime ge '{start_datetime}'"
            elif end_datetime:
                query_params["$filter"] = f"end/dateTime le '{end_datetime}'"

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(
                url, headers=headers, params=query_params, timeout=10
            )
            response.raise_for_status()

            events = response.json().get("value", [])
            logger.info(f"Retrieved {len(events)} events.")
            return {"result": events, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error retrieving events: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    calendar_id: str = "",
    start_datetime: str = "",
    end_datetime: str = "",
) -> Dict[str, Any]:
    try:
        logger.info(
            "Fetching connection credentials.", extra={"connection_id": connection_id}
        )
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )

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

        event_fetcher = OutlookEventFetcher()
        result = event_fetcher.get_events(
            access_token=access_token,
            calendar_id=calendar_id,
            start_datetime=start_datetime,
            end_datetime=end_datetime,
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
