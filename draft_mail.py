from typing import Dict, Any, List
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger

class OutlookDraftSaver:
    @staticmethod
    def prepare_draft(email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Helper method to prepare a single draft message payload"""
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
    def save_draft(
        access_token: str, email_data: Dict[str, Any]
    ) -> Dict[str, Any]:
        try:
            url = "https://graph.microsoft.com/v1.0/me/messages"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # Prepare the draft payload
            draft = OutlookDraftSaver.prepare_draft(email_data)

            # Save the draft
            response = requests.post(url, headers=headers, json=draft, timeout=10)
            response.raise_for_status()

            if response.status_code in [200, 201]:
                logger.info(f"Draft saved successfully for subject: {email_data.get('subject')}")
                return {"status": "success", "draft_id": response.json().get("id"), "error": None}
            else:
                return {
                    "status": "failed",
                    "error": f"Unexpected status code: {response.status_code}"
                }

        except Exception as e:
            logger.error(f"Error saving draft: {e}", extra={"path": os.getenv("WM_JOB_PATH")})
            return {"status": "failed", "error": str(e)}


def main(
    connection_id: str,
    subject: str,
    content: str,
    to_recipients: List[str],
    cc_recipients: List[str] = None,
    bcc_recipients: List[str] = None,
    content_type: str = "HTML",
    importance: str = None,
    attachments: List[Dict[str, Any]] = None,
    custom_headers: List[Dict[str, str]] = None,
    flag: Dict[str, Any] = None
) -> Dict[str, Any]:
    """
    Windmill script to save emails as drafts in Outlook
    
    Args:
        connection_id: The Outlook connection ID from Nango
        subject: Email subject
        content: Email body content
        to_recipients: List of email addresses for TO field
        cc_recipients: Optional list of CC recipients
        bcc_recipients: Optional list of BCC recipients
        content_type: Type of content - HTML or Text
        importance: Email importance level (low, normal, high)
        attachments: Optional list of attachments
        custom_headers: Optional list of custom email headers
        flag: Optional flag settings for the email

    Returns:
        Dict containing the result of the draft saving operation
    """
    # Construct the email data
    email_data = {
        "subject": subject,
        "content": content,
        "contentType": content_type,
        "to": to_recipients
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
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        draft_saver = OutlookDraftSaver()
        return draft_saver.save_draft(access_token=access_token, email_data=email_data)
    except Exception as e:
        error_message = f"Error in draft saving script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
        
