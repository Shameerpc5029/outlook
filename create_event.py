from typing import Optional, Dict, Any, List
from datetime import datetime
import pytz
import tzlocal
import requests
import os
import re
from geopy.geocoders import Nominatim
from geopy.exc import GeopyError
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


def get_user_timezone() -> str:
    """
    Get user's timezone using system settings with UTC fallback.
    """
    try:
        return tzlocal.get_localzone_name()
    except Exception as e:
        logger.warning(
            f"Failed to get system timezone: {e}",
            extra={"path": os.getenv("WM_JOB_PATH")},
        )
        return "UTC"


def validate_email(email: str) -> bool:
    """
    Validate email format using regex.
    """
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(email_regex, email))


class OutlookCalendarCreator:
    DEFAULT_TIMEOUT = 10

    @staticmethod
    def process_attendees(attendees: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Process and validate attendees list.
        """
        processed_attendees = []
        for attendee in attendees:
            try:
                if "email" in attendee and validate_email(attendee["email"]):
                    processed_attendees.append(
                        {
                            "emailAddress": {"address": attendee["email"]},
                            "type": "optional"
                            if attendee.get("optional", False)
                            else "required",
                        }
                    )
                else:
                    logger.warning(f"Invalid or missing email in attendee: {attendee}")
            except Exception as e:
                logger.error(f"Error processing attendee {attendee}: {str(e)}")
        return processed_attendees

    @staticmethod
    def build_event_payload(
        subject: str,
        timezone: str,
        start_time: datetime,
        end_time: datetime,
        body: Optional[str] = None,
        location: Optional[str] = None,
        attendees: Optional[List[Dict[str, Any]]] = None,
        is_online_meeting: bool = False,
    ) -> Dict[str, Any]:
        """
        Build the event payload for the Outlook Calendar API.
        """
        event_data = {
            "subject": subject,
            "start": {"dateTime": start_time.isoformat(), "timeZone": timezone},
            "end": {"dateTime": end_time.isoformat(), "timeZone": timezone},
            "isOnlineMeeting": is_online_meeting,
        }

        if body:
            event_data["body"] = {"contentType": "HTML", "content": body}
        if location:
            event_data["location"] = {"displayName": location}
        if attendees:
            processed_attendees = OutlookCalendarCreator.process_attendees(attendees)
            if processed_attendees:
                event_data["attendees"] = processed_attendees
        return event_data

    def create_event(
        self,
        access_token: str,
        subject: str,
        start_time: datetime,
        end_time: datetime,
        body: Optional[str] = None,
        location: Optional[str] = None,
        attendees: Optional[List[Dict[str, Any]]] = None,
        is_online_meeting: bool = False,
        timeout: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Create a new Outlook Calendar event using Microsoft Graph API.
        """
        try:
            timezone_str = get_user_timezone()
            timezone = pytz.timezone(timezone_str)
            start_time = start_time.astimezone(timezone)
            end_time = end_time.astimezone(timezone)

            # Build event payload
            event_data = self.build_event_payload(
                subject=subject,
                timezone=timezone_str,
                start_time=start_time,
                end_time=end_time,
                body=body,
                location=location,
                attendees=attendees,
                is_online_meeting=is_online_meeting,
            )

            # Make API request
            response = requests.post(
                "https://graph.microsoft.com/v1.0/me/events",
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/json",
                },
                json=event_data,
                timeout=timeout or self.DEFAULT_TIMEOUT,
            )
            response.raise_for_status()

            created_event = response.json()
            logger.info(f"Created event: {created_event.get('id')}")
            return {"result": created_event, "error": None}

        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error creating event: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    start_time: datetime,
    end_time: datetime,
    subject: str,
    body: str = "",
    location: str = "",
    attendees: list = [{"email": "", "optional": True}],
    is_online_meeting: bool = False,
    # provider_config_key: str = "outlook",
) -> Dict[str, Any]:
    """
    Main function to create an Outlook Calendar event.
    """
    try:
        calendar = OutlookCalendarCreator()
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        result = calendar.create_event(
            access_token=access_token,
            subject=subject,
            start_time=start_time,
            end_time=end_time,
            body=body,
            location=location,
            attendees=attendees,
            is_online_meeting=is_online_meeting,
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message)
        return {"result": None, "error": error_message}
