import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookDraftSender:
    @staticmethod
    def send_draft(access_token: str, draft_id: str) -> dict:
        """
        Send a drafted email in Outlook using Microsoft Graph API.

        Args:
            access_token: OAuth token to authenticate with Microsoft Graph API.
            draft_id: The ID of the draft email to send.

        Returns:
            A dictionary containing the status of the operation or an error message.
        """
        try:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_id}/send"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # Make the API call to send the drafted email
            response = requests.post(url, headers=headers, timeout=10)
            response.raise_for_status()

            if response.status_code == 202:  # Accepted
                logger.info(f"Draft email sent successfully: {draft_id}")
                return {"status": "success", "draft_id": draft_id, "error": None}
            else:
                logger.error(
                    f"Failed to send draft: {response.status_code} - {response.text}"
                )
                return {
                    "status": "failed",
                    "draft_id": draft_id,
                    "error": f"Unexpected status code: {response.status_code}",
                }

        except Exception as e:
            logger.error(f"Error sending draft: {e}", extra={"path": os.getenv("WM_JOB_PATH")})
            return {"status": "failed", "draft_id": draft_id, "error": str(e)}


def main(connection_id: str, draft_id: str) -> dict:
    """
    Windmill script to send a drafted email in Outlook.

    Args:
        connection_id: The Outlook connection ID from Nango.
        draft_id: The ID of the draft email to send.

    Returns:
        A dictionary containing the status of the operation or an error message.
    """
    try:
        # Retrieve access token using Nango
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        # Send the draft email
        draft_sender = OutlookDraftSender()
        return draft_sender.send_draft(access_token=access_token, draft_id=draft_id)

    except Exception as e:
        error_message = f"Error in sending draft email script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"status": "failed", "draft_id": draft_id, "error": error_message}
