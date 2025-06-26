from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger

class OutlookDraftDeleter:
    @staticmethod
    def delete_draft(access_token: str, draft_id: str) -> Dict[str, Any]:
        """
        Delete a draft email from Outlook using Microsoft Graph API.
        
        Args:
            access_token: OAuth token to authenticate with Microsoft Graph API.
            draft_id: The ID of the draft email to delete.
        
        Returns:
            Dict containing the status of the deletion or an error message.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            
            # Make the API request to delete the draft
            response = requests.delete(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            # Parse the response
            if response.status_code == 204:  # No Content indicates success
                logger.info(f"Draft email deleted successfully: {draft_id}")
                return {"status": "success", "draft_id": draft_id, "error": None}
            else:
                logger.error(f"Failed to delete draft: {response.status_code} - {response.text}")
                return {"status": "failed", "draft_id": draft_id, "error": f"Unexpected status code: {response.status_code}"}
        
        except Exception as e:
            logger.error(f"Error deleting draft: {e}", extra={"path": os.getenv("WM_JOB_PATH")})
            return {"status": "failed", "draft_id": draft_id, "error": str(e)}

def main(connection_id: str, draft_id: str) -> Dict[str, Any]:
    """
    Windmill script to delete a draft email from Outlook.
    
    Args:
        connection_id: The Outlook connection ID from Nango.
        draft_id: The ID of the draft email to delete.
    
    Returns:
        Dict containing the status of the deletion or an error message.
    """
    try:
        # Retrieve access token using Nango
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        
        # Delete the draft email
        draft_deleter = OutlookDraftDeleter()
        return draft_deleter.delete_draft(access_token=access_token, draft_id=draft_id)
    
    except Exception as e:
        error_message = f"Error in deleting draft email script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"status": "failed", "draft_id": draft_id, "error": error_message}
