from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger

class OutlookDraftUpdater:
    @staticmethod
    def update_draft(access_token: str, draft_id: str, updates: Dict[str, Any]) -> Dict[str, Any]:
        """
        Update a draft email in Outlook using Microsoft Graph API.
        
        Args:
            access_token: OAuth token to authenticate with Microsoft Graph API.
            draft_id: The ID of the draft email to update.
            updates: The fields and values to update in the draft.
        
        Returns:
            Dict containing the status of the update operation or an error message.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            
            # Make the API request to update the draft
            response = requests.patch(url, headers=headers, json=updates, timeout=10)
            response.raise_for_status()
            
            # Parse the response
            if response.status_code == 200:  # Success
                logger.info(f"Draft email updated successfully: {draft_id}")
                return {"status": "success", "draft_id": draft_id, "error": None}
            else:
                logger.error(f"Failed to update draft: {response.status_code} - {response.text}")
                return {"status": "failed", "draft_id": draft_id, "error": f"Unexpected status code: {response.status_code}"}
        
        except Exception as e:
            logger.error(f"Error updating draft: {e}", extra={"path": os.getenv("WM_JOB_PATH")})
            return {"status": "failed", "draft_id": draft_id, "error": str(e)}

def main(
    connection_id: str,
    draft_id: str,
    subject: str = None,
    content: str = None,
    content_type: str = "HTML",
    to_recipients: list = None,
    cc_recipients: list = None,
    bcc_recipients: list = None,
    importance: str = None,
    attachments: list = None
) -> Dict[str, Any]:
    """
    Windmill script to update a draft email in Outlook.
    
    Args:
        connection_id: The Outlook connection ID from Nango.
        draft_id: The ID of the draft email to update.
        subject: Optional new subject for the draft.
        content: Optional content of the draft body.
        content_type: Optional content type of the body (e.g., "HTML" or "Text").
        to_recipients: Optional list of email addresses for TO field.
        cc_recipients: Optional list of CC recipients.
        bcc_recipients: Optional list of BCC recipients.
        importance: Optional importance level (low, normal, high).
        attachments: Optional list of updated attachments.
    
    Returns:
        Dict containing the status of the update operation or an error message.
    """
    try:
        # Construct the updates payload
        updates = {}
        
        if subject:
            updates["subject"] = subject
        
        if content:
            updates["body"] = {
                "contentType": content_type,
                "content": content
            }
        
        if to_recipients:
            updates["toRecipients"] = [
                {"emailAddress": {"address": recipient}} for recipient in to_recipients
            ]
        
        if cc_recipients:
            updates["ccRecipients"] = [
                {"emailAddress": {"address": recipient}} for recipient in cc_recipients
            ]
        
        if bcc_recipients:
            updates["bccRecipients"] = [
                {"emailAddress": {"address": recipient}} for recipient in bcc_recipients
            ]
        
        if importance:
            updates["importance"] = importance
        
        if attachments:
            updates["attachments"] = [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name", ""),
                    "contentType": attachment.get("contentType", ""),
                    "contentBytes": attachment.get("contentBytes", "")
                } for attachment in attachments
            ]
        
        # Retrieve access token using Nango
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        
        # Update the draft email
        draft_updater = OutlookDraftUpdater()
        return draft_updater.update_draft(
            access_token=access_token, draft_id=draft_id, updates=updates
        )
    
    except Exception as e:
        error_message = f"Error in updating draft email script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"status": "failed", "draft_id": draft_id, "error": error_message}
