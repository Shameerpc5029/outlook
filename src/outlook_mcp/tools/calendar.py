"""Calendar management tools for Outlook MCP Server"""
from typing import Dict, Any, Optional, List
import requests
from ..connection import get_access_token


def get_all_calendars() -> Dict[str, Any]:
    """
    Fetch all calendars and return only id, name, and owner details.
    """
    try:
        access_token = get_access_token()
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

        print(f"Fetched {len(filtered_calendars)} calendars.")
        return {"result": filtered_calendars, "error": None}
        
    except requests.exceptions.RequestException as e:
        error_message = f"API request failed: {e}"
        print(error_message)
        return {"result": None, "error": error_message}
    except Exception as e:
        error_message = f"Error fetching calendars: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_calendar_details(calendar_id: str) -> Dict[str, Any]:
    """Get details of a specific calendar"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        calendar = response.json()
        print(f"Retrieved calendar details for: {calendar.get('name')}")
        return {"result": calendar, "error": None}
        
    except Exception as e:
        error_message = f"Error getting calendar details: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def create_calendar(
    name: str,
    color: str = "auto"
) -> Dict[str, Any]:
    """Create a new calendar"""
    try:
        access_token = get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/calendars"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        calendar_data = {
            "name": name,
            "color": color
        }

        response = requests.post(url, headers=headers, json=calendar_data, timeout=10)
        response.raise_for_status()

        calendar = response.json()
        print(f"Created calendar: {calendar.get('name')}")
        return {"result": calendar, "error": None}
        
    except Exception as e:
        error_message = f"Error creating calendar: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def update_calendar(
    calendar_id: str,
    name: Optional[str] = None,
    color: Optional[str] = None
) -> Dict[str, Any]:
    """Update an existing calendar"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        update_data = {}
        if name:
            update_data["name"] = name
        if color:
            update_data["color"] = color

        response = requests.patch(url, headers=headers, json=update_data, timeout=10)
        response.raise_for_status()

        updated_calendar = response.json()
        print(f"Updated calendar: {calendar_id}")
        return {"result": updated_calendar, "error": None}
        
    except Exception as e:
        error_message = f"Error updating calendar: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def delete_calendar(calendar_id: str) -> Dict[str, Any]:
    """Delete a calendar"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.delete(url, headers=headers, timeout=10)
        response.raise_for_status()

        print(f"Deleted calendar: {calendar_id}")
        return {"result": "Calendar deleted successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error deleting calendar: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_all_events(calendar_id: Optional[str] = None) -> Dict[str, Any]:
    """Get all events from a specific calendar or default calendar"""
    try:
        access_token = get_access_token()
        
        if calendar_id:
            url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events"
        else:
            url = "https://graph.microsoft.com/v1.0/me/events"
            
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        events = response.json().get("value", [])
        filtered_events = [
            {
                "id": event.get("id"),
                "subject": event.get("subject"),
                "start": event.get("start"),
                "end": event.get("end"),
                "organizer": event.get("organizer", {}).get("emailAddress", {}).get("address"),
                "location": event.get("location", {}).get("displayName"),
                "attendees": [
                    attendee.get("emailAddress", {}).get("address")
                    for attendee in event.get("attendees", [])
                ]
            }
            for event in events
        ]

        print(f"Retrieved {len(filtered_events)} events")
        return {"result": filtered_events, "error": None}
        
    except Exception as e:
        error_message = f"Error getting events: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_event_details(event_id: str) -> Dict[str, Any]:
    """Get details of a specific event"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        event = response.json()
        print(f"Retrieved event details for: {event.get('subject')}")
        return {"result": event, "error": None}
        
    except Exception as e:
        error_message = f"Error getting event details: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def create_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    start_timezone: str = "UTC",
    end_timezone: str = "UTC",
    body_content: str = "",
    body_content_type: str = "HTML",
    location: Optional[str] = None,
    attendees: Optional[List[str]] = None,
    calendar_id: Optional[str] = None
) -> Dict[str, Any]:
    """Create a new event"""
    try:
        access_token = get_access_token()
        
        if calendar_id:
            url = f"https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events"
        else:
            url = "https://graph.microsoft.com/v1.0/me/events"
            
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        event_data = {
            "subject": subject,
            "start": {
                "dateTime": start_datetime,
                "timeZone": start_timezone
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": end_timezone
            },
            "body": {
                "contentType": body_content_type,
                "content": body_content
            }
        }

        if location:
            event_data["location"] = {"displayName": location}

        if attendees:
            event_data["attendees"] = [
                {
                    "emailAddress": {"address": attendee},
                    "type": "required"
                }
                for attendee in attendees
            ]

        response = requests.post(url, headers=headers, json=event_data, timeout=10)
        response.raise_for_status()

        event = response.json()
        print(f"Created event: {event.get('subject')}")
        return {"result": event, "error": None}
        
    except Exception as e:
        error_message = f"Error creating event: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def delete_event(event_id: str) -> Dict[str, Any]:
    """Delete an event"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.delete(url, headers=headers, timeout=10)
        response.raise_for_status()

        print(f"Deleted event: {event_id}")
        return {"result": "Event deleted successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error deleting event: {e}"
        print(error_message)
        return {"result": None, "error": error_message}
