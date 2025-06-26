from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger

class OutlookDraftRetriever:
    @staticmethod
    def get_draft(access_token: str, draft_id: str) -> Dict[str, Any]:
        """
        Retrieve a draft email from Outlook using Microsoft Graph API.
        
        Args:
            access_token: OAuth token to authenticate with Microsoft Graph API.
            draft_id: The ID of the draft email to retrieve.
        
        Returns:
            Dict containing the draft details or an error message.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }
            
            # Make the API request to get the draft
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            # Parse the response
            if response.status_code == 200:
                draft = response.json()
                logger.info(f"Draft email retrieved successfully: {draft_id}")
                return {"status": "success", "draft": draft, "error": None}
            else:
                logger.error(f"Failed to retrieve draft: {response.status_code} - {response.text}")
                return {"status": "failed", "draft": None, "error": f"Unexpected status code: {response.status_code}"}
        
        except Exception as e:
            logger.error(f"Error retrieving draft: {e}", extra={"path": os.getenv("WM_JOB_PATH")})
            return {"status": "failed", "draft": None, "error": str(e)}

def main(connection_id: str, draft_id: str) -> Dict[str, Any]:
    """
    Windmill script to retrieve a draft email from Outlook.
    
    Args:
        connection_id: The Outlook connection ID from Nango.
        draft_id: The ID of the draft email to retrieve.
    
    Returns:
        Dict containing the draft details or an error message.
    """
    try:
        # Retrieve access token using Nango
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        
        # Retrieve the draft email
        draft_retriever = OutlookDraftRetriever()
        return draft_retriever.get_draft(access_token=access_token, draft_id=draft_id)
    
    except Exception as e:
        error_message = f"Error in retrieving draft email script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"status": "failed", "draft": None, "error": error_message}
