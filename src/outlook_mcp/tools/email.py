"""Email management tools for Outlook MCP Server"""
from typing import Dict, Any, List, Optional
import requests
from ..connection import get_access_token


class OutlookEmailSender:
    @staticmethod
    def prepare_message(email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Helper method to prepare a single message payload"""
        message = {
            "subject": email_data.get("subject", ""),
            "body": {
                "contentType": email_data.get("contentType", "HTML"),
                "content": email_data.get("content", "")
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                } for recipient in email_data.get("to", [])
            ]
        }
        
        # Add CC recipients if provided
        if "cc" in email_data:
            message["ccRecipients"] = [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                } for recipient in email_data["cc"]
            ]
        
        # Add BCC recipients if provided
        if "bcc" in email_data:
            message["bccRecipients"] = [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                } for recipient in email_data["bcc"]
            ]
        
        # Add attachments if provided
        if "attachments" in email_data:
            message["attachments"] = [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name", ""),
                    "contentType": attachment.get("contentType", ""),
                    "contentBytes": attachment.get("contentBytes", "")
                } for attachment in email_data["attachments"]
            ]

        # Add custom headers if provided
        if "internetMessageHeaders" in email_data:
            message["internetMessageHeaders"] = email_data["internetMessageHeaders"]

        # Add importance if provided
        if "importance" in email_data:
            message["importance"] = email_data["importance"]

        # Add flag if provided
        if "flag" in email_data:
            message["flag"] = email_data["flag"]
            
        return message

    @staticmethod
    def send_emails(emails_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Send multiple emails"""
        try:
            access_token = get_access_token()
            url = "https://graph.microsoft.com/v1.0/me/sendMail"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            
            results = []
            for email_data in emails_data:
                try:
                    # Prepare the message payload
                    message = OutlookEmailSender.prepare_message(email_data)
                    payload = {
                        "message": message,
                        "saveToSentItems": email_data.get("saveToSentItems", True)
                    }

                    # Send the email
                    response = requests.post(
                        url, headers=headers, json=payload, timeout=10
                    )
                    response.raise_for_status()
                    
                    # Track the result
                    if response.status_code == 202:
                        print(f"Email sent successfully to {', '.join(email_data.get('to', []))}")
                        results.append({
                            "status": "success",
                            "recipients": email_data.get("to", []),
                            "error": None
                        })
                    else:
                        results.append({
                            "status": "failed",
                            "recipients": email_data.get("to", []),
                            "error": f"Unexpected status code: {response.status_code}"
                        })
                        
                except Exception as e:
                    print(f"Error sending individual email: {e}")
                    results.append({
                        "status": "failed",
                        "recipients": email_data.get("to", []),
                        "error": str(e)
                    })
            
            return {"result": results, "error": None}
            
        except Exception as e:
            error_message = f"Error in batch email sending: {e}"
            print(error_message)
            return {"result": None, "error": error_message}


def send_email(
    subject: str,
    content: str,
    to_recipients: List[str],
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    content_type: str = "HTML",
    save_to_sent: bool = True,
    importance: Optional[str] = None,
    attachments: Optional[List[Dict[str, Any]]] = None,
    custom_headers: Optional[List[Dict[str, str]]] = None,
    flag: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Send email via Outlook with all available options
    
    Args:
        subject: Email subject
        content: Email body content
        to_recipients: List of email addresses for TO field
        cc_recipients: Optional list of CC recipients
        bcc_recipients: Optional list of BCC recipients
        content_type: Type of content - HTML or Text
        save_to_sent: Whether to save in Sent Items folder
        importance: Email importance level (low, normal, high)
        attachments: Optional list of attachments
        custom_headers: Optional list of custom email headers
        flag: Optional flag settings for the email
        
    Returns:
        Dict containing the result of the email sending operation
    """
    # Construct the email data
    email_data = {
        "subject": subject,
        "content": content,
        "contentType": content_type,
        "to": to_recipients,
        "saveToSentItems": save_to_sent
    }
    
    # Add optional CC recipients
    if cc_recipients:
        email_data["cc"] = cc_recipients
        
    # Add optional BCC recipients
    if bcc_recipients:
        email_data["bcc"] = bcc_recipients
        
    # Add importance if specified
    if importance:
        email_data["importance"] = importance
        
    # Add attachments if provided
    if attachments:
        email_data["attachments"] = attachments
        
    # Add custom headers if provided
    if custom_headers:
        email_data["internetMessageHeaders"] = custom_headers
        
    # Add flag if provided
    if flag:
        email_data["flag"] = flag

    try:
        email_sender = OutlookEmailSender()
        return email_sender.send_emails(emails_data=[email_data])
    except Exception as e:
        error_message = f"Error in email sending: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def create_draft_email(
    subject: str,
    content: str,
    to_recipients: List[str],
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    content_type: str = "HTML",
    importance: Optional[str] = None,
    attachments: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """Create a draft email"""
    try:
        access_token = get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/messages"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        
        # Prepare the message payload
        message = {
            "subject": subject,
            "body": {
                "contentType": content_type,
                "content": content
            },
            "toRecipients": [
                {"emailAddress": {"address": recipient}}
                for recipient in to_recipients
            ]
        }
        
        # Add optional fields
        if cc_recipients:
            message["ccRecipients"] = [
                {"emailAddress": {"address": recipient}}
                for recipient in cc_recipients
            ]
        
        if bcc_recipients:
            message["bccRecipients"] = [
                {"emailAddress": {"address": recipient}}
                for recipient in bcc_recipients
            ]
            
        if importance:
            message["importance"] = importance
            
        if attachments:
            message["attachments"] = attachments

        response = requests.post(url, headers=headers, json=message, timeout=10)
        response.raise_for_status()
        
        draft = response.json()
        print(f"Draft created successfully with ID: {draft.get('id')}")
        return {"result": draft, "error": None}
        
    except Exception as e:
        error_message = f"Error creating draft: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def send_draft_email(draft_id: str) -> Dict[str, Any]:
    """Send a draft email"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}/send"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.post(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        print(f"Draft {draft_id} sent successfully")
        return {"result": "Draft sent successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error sending draft: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def get_draft_emails() -> Dict[str, Any]:
    """Get all draft emails"""
    try:
        access_token = get_access_token()
        url = "https://graph.microsoft.com/v1.0/me/mailFolders/drafts/messages"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        drafts = response.json().get("value", [])
        filtered_drafts = [
            {
                "id": draft.get("id"),
                "subject": draft.get("subject"),
                "bodyPreview": draft.get("bodyPreview"),
                "createdDateTime": draft.get("createdDateTime"),
                "lastModifiedDateTime": draft.get("lastModifiedDateTime"),
                "toRecipients": [
                    recipient.get("emailAddress", {}).get("address")
                    for recipient in draft.get("toRecipients", [])
                ]
            }
            for draft in drafts
        ]
        
        print(f"Retrieved {len(filtered_drafts)} draft emails")
        return {"result": filtered_drafts, "error": None}
        
    except Exception as e:
        error_message = f"Error getting drafts: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def delete_draft_email(draft_id: str) -> Dict[str, Any]:
    """Delete a draft email"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        response = requests.delete(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        print(f"Draft {draft_id} deleted successfully")
        return {"result": "Draft deleted successfully", "error": None}
        
    except Exception as e:
        error_message = f"Error deleting draft: {e}"
        print(error_message)
        return {"result": None, "error": error_message}


def update_draft_email(
    draft_id: str,
    subject: Optional[str] = None,
    content: Optional[str] = None,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    content_type: str = "HTML",
    importance: Optional[str] = None
) -> Dict[str, Any]:
    """Update a draft email"""
    try:
        access_token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        
        # Build update payload
        update_data = {}
        
        if subject:
            update_data["subject"] = subject
            
        if content:
            update_data["body"] = {
                "contentType": content_type,
                "content": content
            }
            
        if to_recipients:
            update_data["toRecipients"] = [
                {"emailAddress": {"address": recipient}}
                for recipient in to_recipients
            ]
            
        if cc_recipients:
            update_data["ccRecipients"] = [
                {"emailAddress": {"address": recipient}}
                for recipient in cc_recipients
            ]
            
        if bcc_recipients:
            update_data["bccRecipients"] = [
                {"emailAddress": {"address": recipient}}
                for recipient in bcc_recipients
            ]
            
        if importance:
            update_data["importance"] = importance

        response = requests.patch(url, headers=headers, json=update_data, timeout=10)
        response.raise_for_status()
        
        updated_draft = response.json()
        print(f"Draft {draft_id} updated successfully")
        return {"result": updated_draft, "error": None}
        
    except Exception as e:
        error_message = f"Error updating draft: {e}"
        print(error_message)
        return {"result": None, "error": error_message}
