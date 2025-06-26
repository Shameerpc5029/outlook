from typing import Optional, Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookCalendarManager:
    @staticmethod
    def validate_calendar_name(name: str) -> None:
        """
        Validate the calendar name.
        """
        if not name or not isinstance(name, str):
            raise ValueError("Calendar name must be a non-empty string")
        if len(name) > 255:  # Microsoft's limit
            raise ValueError("Calendar name must be less than 255 characters")

    @staticmethod
    def create_calendar(
        access_token: str,
        name: str,
        color: Optional[str] = None,
        is_default: bool = False,
        is_removable: bool = True,
    ) -> Dict[str, Any]:
        """
        Create a new calendar in Outlook using Microsoft Graph API.

        Args:
            access_token (str): The OAuth access token
            name (str): Name of the calendar
            color (str, optional): Calendar color (e.g., "lightBlue", "lightGreen")
            is_default (bool): Whether this should be the default calendar
            is_removable (bool): Whether the calendar can be deleted
        """
        try:
            # Validate the calendar name
            OutlookCalendarManager.validate_calendar_name(name)

            # Prepare payload for the new calendar
            payload = {
                "name": name,
                "isDefaultCalendar": is_default,
                "isRemovable": is_removable,
            }

            # Add optional color if provided
            if color:
                payload["color"] = color

            # Make the API request to create the calendar
            response = requests.post(
                "https://graph.microsoft.com/v1.0/me/calendars",
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/json",
                    "Prefer": 'outlook.timezone="UTC"',
                },
                json=payload,
                timeout=10,
            )

            if response.status_code == 401:
                logger.error(
                    "Authentication failed: Invalid or expired access token",
                    extra={"path": os.getenv("WM_JOB_PATH")},
                )
                return {"result": None, "error": "Authentication failed"}

            response.raise_for_status()

            # Return the created calendar details
            created_calendar = response.json()
            logger.info(
                f"Created Outlook calendar: {created_calendar.get('id')}",
                extra={
                    "path": os.getenv("WM_JOB_PATH"),
                    "calendar_id": created_calendar.get("id"),
                    "name": created_calendar.get("name"),
                    "is_default": is_default,
                    "color": color,
                },
            )
            return {"result": created_calendar, "error": None}

        except requests.exceptions.RequestException as e:
            status_code = (
                getattr(e.response, "status_code", None)
                if hasattr(e, "response")
                else None
            )
            error_message = f"Outlook API request failed: {str(e)}"

            logger.error(
                error_message,
                extra={
                    "path": os.getenv("WM_JOB_PATH"),
                    "status_code": status_code,
                    "calendar_name": name,
                },
            )
            return {"result": None, "error": error_message}

        except ValueError as e:
            error_message = str(e)
            logger.error(
                f"Validation error: {error_message}",
                extra={"path": os.getenv("WM_JOB_PATH")},
            )
            return {"result": None, "error": error_message}

        except Exception as e:
            error_message = str(e)
            logger.error(
                f"Error creating calendar: {error_message}",
                extra={"path": os.getenv("WM_JOB_PATH"), "calendar_name": name},
            )
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    calendar_name: str,
    # color: Optional[str] = None,
    # is_default: bool = False,
    # is_removable: bool = True,
) -> Dict[str, Any]:
    """
    Main function to create a new calendar in Outlook.

    Args:
        connection_id (str): Nango connection ID
        calendar_name (str): Name of the calendar
        color (str, optional): Calendar color
        is_default (bool): Whether this should be the default calendar
        is_removable (bool): Whether the calendar can be deleted
    """
    try:
        # Get credentials from Nango
        credentials = get_connection_credentials(
            id=connection_id,
            providerConfigKey="outlook",
        )
        if not credentials or "credentials" not in credentials:
            raise ValueError("Failed to retrieve credentials from Nango")

        # Extract the access token
        access_token = credentials["credentials"]["access_token"]

        # Create the calendar
        calendar_manager = OutlookCalendarManager()
        result = calendar_manager.create_calendar(
            access_token=access_token,
            name=calendar_name,
            # color=color,
            # is_default=is_default,
            # is_removable=is_removable,
        )

        return result

    except Exception as e:
        error_message = str(e)
        logger.error(
            f"Error in main function: {error_message}",
            extra={"path": os.getenv("WM_JOB_PATH"), "connection_id": connection_id},
        )
        return {"result": None, "error": error_message}
