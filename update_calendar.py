from typing import Optional, Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookCalendarUpdater:
    @staticmethod
    def update_calendar_by_id(
        access_token: str,
        calendar_id: str,
        name: str = "",
        color: str = "",
        is_default: bool = False,
    ) -> Dict[str, Any]:
        """
        Update a calendar in Microsoft Outlook using Microsoft Graph API.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # Build the payload for updating the calendar
            payload = {}
            if name:
                payload["name"] = name
            if color:
                payload["color"] = color
            if is_default is not None:
                payload["isDefaultCalendar"] = is_default

            if not payload:
                raise ValueError("No properties provided to update the calendar.")

            logger.info(
                "Sending PATCH request to update calendar.",
                extra={"url": url, "payload": payload},
            )

            response = requests.patch(url, headers=headers, json=payload, timeout=10)
            response.raise_for_status()

            updated_calendar = response.json()
            logger.info(
                f"Successfully updated calendar with ID: {calendar_id}",
                extra={"updated_calendar": updated_calendar},
            )
            return {"result": updated_calendar, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error updating calendar: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    calendar_id: str,
    name: str = "",
    color: str = "",
    is_default: bool = False,
) -> Dict[str, Any]:
    """
    Supported Color Values:
    Auto
    LightBlue
    LightGreen
    LightOrange
    LightGray
    LightYellow
    LightTeal
    LightPink
    LightBrown
    LightRed
    MaxColor (Special case, not commonly used)
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

        calendar_updater = OutlookCalendarUpdater()
        result = calendar_updater.update_calendar_by_id(
            access_token=access_token,
            calendar_id=calendar_id,
            name=name,
            color=color,
            is_default=is_default,
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
