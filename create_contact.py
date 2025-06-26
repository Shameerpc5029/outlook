from typing import Optional, Dict, Any, List
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookContactCreator:
    @staticmethod
    def build_contact_payload(
        given_name: str,
        surname: Optional[str] = None,
        email_addresses: Optional[str] = None,  # Changed to str
        business_phones: Optional[str] = None,  # Changed to str
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
        self, access_token: str, contact_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Create a new contact using Microsoft Graph API.

        Args:
            access_token: The OAuth access token for Microsoft Graph API
            contact_data: The contact data payload

        Returns:
            Dictionary containing either the created contact or error information
        """
        try:
            url = "https://graph.microsoft.com/v1.0/me/contacts"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.post(
                url, headers=headers, json=contact_data, timeout=10
            )
            response.raise_for_status()

            contact = response.json()
            logger.info(f"Created contact: {contact.get('id')}")
            return {"result": contact, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error creating contact: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(
    connection_id: str,
    given_name: str,
    surname: str = "",
    email_addresses: str = "",  # Changed to str
    business_phones: str = "",  # Changed to str
    mobile_phone: str = "",
    job_title: str = "",
    company_name: str = "",
    department: str = "",
    office_location: str = "",
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

        logger.info(
            "Sending request to create contact.", extra={"contact_data": contact_data}
        )
        result = contact_creator.create_contact(
            access_token=access_token, contact_data=contact_data
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
