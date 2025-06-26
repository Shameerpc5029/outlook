from typing import Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookEventDetailsFetcher:
    @staticmethod
    def get_event_details(access_token: str, event_id: str) -> Dict[str, Any]:
        """
        Fetch detailed information for a specific Outlook calendar event.
        
        :param access_token: OAuth2 access token for Microsoft Graph API
        :param event_id: ID of the event to fetch details for
        :return: Dictionary containing the event details or error
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
                "Prefer": "outlook.timezone=\"UTC\""
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            event = response.json()
            
            # Extract and structure all available event details
            event_details = {
                # Basic information
                "id": event.get("id"),
                "subject": event.get("subject"),
                "body": {
                    "content": event.get("body", {}).get("content"),
                    "contentType": event.get("body", {}).get("contentType")
                },
                
                # Timing information
                "start": {
                    "dateTime": event.get("start", {}).get("dateTime"),
                    "timeZone": event.get("start", {}).get("timeZone")
                },
                "end": {
                    "dateTime": event.get("end", {}).get("dateTime"),
                    "timeZone": event.get("end", {}).get("timeZone")
                },
                "isAllDay": event.get("isAllDay", False),
                
                # Location details
                "location": {
                    "displayName": event.get("location", {}).get("displayName"),
                    "address": event.get("location", {}).get("address"),
                    "coordinates": event.get("location", {}).get("coordinates")
                },
                
                # Participants
                "organizer": {
                    "name": event.get("organizer", {}).get("emailAddress", {}).get("name"),
                    "email": event.get("organizer", {}).get("emailAddress", {}).get("address")
                },
                "attendees": [
                    {
                        "name": attendee.get("emailAddress", {}).get("name"),
                        "email": attendee.get("emailAddress", {}).get("address"),
                        "type": attendee.get("type"),
                        "response": attendee.get("status", {}).get("response"),
                        "time": attendee.get("status", {}).get("time")
                    }
                    for attendee in event.get("attendees", [])
                ],
                
                # Online meeting details
                "isOnlineMeeting": event.get("isOnlineMeeting", False),
                "onlineMeeting": {
                    "joinUrl": event.get("onlineMeeting", {}).get("joinUrl"),
                    "conferenceId": event.get("onlineMeeting", {}).get("conferenceId"),
                    "provider": event.get("onlineMeeting", {}).get("provider"),
                    "tollNumber": event.get("onlineMeeting", {}).get("tollNumber")
                } if event.get("isOnlineMeeting") else None,
                
                # Additional properties
                "importance": event.get("importance"),
                "sensitivity": event.get("sensitivity"),
                "showAs": event.get("showAs"),
                "categories": event.get("categories", []),
                
                # Recurrence and reminders
                "recurrence": event.get("recurrence"),
                "reminderMinutesBeforeStart": event.get("reminderMinutesBeforeStart"),
                "isReminderOn": event.get("isReminderOn"),
                
                # Status information
                "responseStatus": {
                    "response": event.get("responseStatus", {}).get("response"),
                    "time": event.get("responseStatus", {}).get("time")
                },
                "isCancelled": event.get("isCancelled", False),
                
                # Timestamps
                "createdDateTime": event.get("createdDateTime"),
                "lastModifiedDateTime": event.get("lastModifiedDateTime"),
                
                # Additional metadata
                "changeKey": event.get("changeKey"),
                "seriesMasterId": event.get("seriesMasterId"),
                "type": event.get("type")
            }
            
            logger.info(f"Successfully fetched details for event {event_id}")
            return {"result": event_details, "error": None}
            
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed for event {event_id}: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching event details for {event_id}: {e}"
            logger.error(error_message)
            return {"result": None, "error": error_message}


def main(connection_id: str, event_id: str) -> Dict[str, Any]:
    """
    Main function to fetch detailed information for a specific Outlook calendar event.
    
    :param connection_id: Nango connection ID for the Outlook account
    :param event_id: ID of the event to fetch details for
    :return: Dictionary containing the result or error
    """
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        
        event_fetcher = OutlookEventDetailsFetcher()
        result = event_fetcher.get_event_details(
            access_token=access_token,
            event_id=event_id
        )
        return result
        
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message)
        return {"result": None, "error": error_message}