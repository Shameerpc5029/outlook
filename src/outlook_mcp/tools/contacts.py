"""Contact management tools for Outlook MCP Server"""
from typing import Optional, Dict, Any
import requests
from ..connection import get_access_token


class OutlookContactCreator:
    @staticmethod
    def build_contact_payload(
        given_name: str,
        surname: Optional[str] = None,
        email_addresses: Optional[str] = None,
        business_phones: Optional[str] = None,
        mobile_phone: Optional[str] = None,
        job_title: Optional[str] = None,
        company_name: Optional[str] = None,
        department: Optional[str] = None,
        office_location: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Build the contact payload for the Microsoft Graph API.

        Args:
            given_name: First name of the contact
            surname: Last name of the contact
            email_addresses: Comma-separated string of email addresses
            business_phones: Comma-separated string of business phone numbers
            mobile_phone: Mobile phone number
            job_title: Job title
            company_name: Company name
            department: Department name
            office_location: Office location

        Returns:
            Dictionary matching the Microsoft Graph API contact schema
        """
        payload: Dict[str, Any] = {"givenName": given_name}

        if surname:
            payload["surname"] = surname
        if email_addresses:
            # Convert comma-separated string to list of email objects
            email_list = [e.strip() for e in email_addresses.split(",")]
            payload["emailAddresses"] = [{"address": email} for email in email_list]
        if business_phones:
            # Convert comma-separated string to list
            payload["businessPhones"] = [p.strip() for p in business_phones.split(",")]
        if mobile_phone:
            payload["mobilePhone"] = mobile_phone
        if job_title:
            payload["jobTitle"] = job_title
        if company_name:
            payload["companyName"] = company_name
        if department:
            payload["department"] = department
        if office_location:
            payload["officeLocation"] = office_location

        return payload


def create_contact(
    given_name: str,
    surname: str = "",
    email_addresses: str = "",
    business_phones: str = "",
    mobile_phone: str = "",
    job_title: str = "",
    company_name: str = "",
    department: str = "",
    office_location: str = "",
) -> Dict[str, Any]:
    """Create a new contact in Outlook"""
    try:
        access_token = get_access_token()
        
        contact_creator = OutlookContactCreator()
        contact_data = contact_creator.build_contact_payload(
            given_name=given_name,
            surname=surname,
            email_addresses=email_addresses,
            business_phones=business_phones,
            mobile_phone=mobile_phone,
            job_title=job_title,
            company_name=company_name,
            department=department,
            office_location=office_location,
        )

        url = "https://graph.microsoft.com/v1.0/me/contacts"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.post(url, headers=headers, json=contact_data, timeout=10)
        response.raise_for_status()

        contact = response.json()
        print(f"Created contact: {contact.get('id')}")
        return {"result": contact, "error": None}
        
    except requests.exceptions.RequestException as e:
        error_message = f"API request failed: {e}"
        print(error_message)
        return {"result": None, "error": error_message}
    except Exception as e:
        error_message = f"Error creating contact: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_all_contacts() -> Dict[str, Any]:
    """Get all contacts from Outlook"""
    try:
        access_token = get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/contacts"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        contacts = response.json().get("value", [])
        filtered_contacts = [
            {
                "id": contact.get("id"),
                "displayName": contact.get("displayName"),
                "givenName": contact.get("givenName"),
                "surname": contact.get("surname"),
                "emailAddresses": contact.get("emailAddresses", []),
                "businessPhones": contact.get("businessPhones", []),
                "mobilePhone": contact.get("mobilePhone"),
                "jobTitle": contact.get("jobTitle"),
                "companyName": contact.get("companyName"),
            }
            for contact in contacts
        ]

        print(f"Retrieved {len(filtered_contacts)} contacts")
        return {"result": filtered_contacts, "error": None}
        
    except Exception as e:
        error_message = f"Error getting contacts: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_contact_details(contact_id: str) -> Dict[str, Any]:
    """Get details of a specific contact"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        contact = response.json()
        print(f"Retrieved contact details for: {contact.get('displayName')}")
        return {"result": contact, "error": None}
        
    except Exception as e:
        error_message = f"Error getting contact details: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def update_contact(
    contact_id: str,
    given_name: Optional[str] = None,
    surname: Optional[str] = None,
    email_addresses: Optional[str] = None,
    business_phones: Optional[str] = None,
    mobile_phone: Optional[str] = None,
    job_title: Optional[str] = None,
    company_name: Optional[str] = None,
    department: Optional[str] = None,
    office_location: Optional[str] = None,
) -> Dict[str, Any]:
    """Update an existing contact"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        # Build update payload
        update_data = {}
        
        if given_name:
            update_data["givenName"] = given_name
        if surname:
            update_data["surname"] = surname
        if email_addresses:
            email_list = [e.strip() for e in email_addresses.split(",")]
            update_data["emailAddresses"] = [{"address": email} for email in email_list]
        if business_phones:
            update_data["businessPhones"] = [p.strip() for p in business_phones.split(",")]
        if mobile_phone:
            update_data["mobilePhone"] = mobile_phone
        if job_title:
            update_data["jobTitle"] = job_title
        if company_name:
            update_data["companyName"] = company_name
        if department:
            update_data["department"] = department
        if office_location:
            update_data["officeLocation"] = office_location

        response = requests.patch(url, headers=headers, json=update_data, timeout=10)
        response.raise_for_status()

        updated_contact = response.json()
        print(f"Updated contact: {contact_id}")
        return {"result": updated_contact, "error": None}
        
    except Exception as e:
        error_message = f"Error updating contact: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def delete_contact(contact_id: str) -> Dict[str, Any]:
    """Delete a contact"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.delete(url, headers=headers, timeout=10)
        response.raise_for_status()

        print(f"Deleted contact: {contact_id}")
        return {"result": "Contact deleted successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error deleting contact: {e}"
        print(error_message)
        return {"result": None, "error": error_message}
