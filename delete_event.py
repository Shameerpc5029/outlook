from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class DeleteOutlookEvent:
    @staticmethod
    def delete_event_by_id(access_token: str, event_id: str) -> Dict[str, Any]:
        try:
            url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.delete(url, headers=headers, timeout=10)
            response.raise_for_status()

            logger.info(f"Successfully deleted event with ID: {event_id}")
            return {"result": "Event deleted successfully", "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error deleting event: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(connection_id: str, event_id: str) -> Dict[str, Any]:
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

        event_deleter = DeleteOutlookEvent()
        result = event_deleter.delete_event_by_id(
            access_token=access_token, event_id=event_id
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
