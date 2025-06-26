from typing import Dict, Any
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger


class OutlookContactFetcher:
    @staticmethod
    def get_contact_by_id(access_token: str, contact_id: str) -> Dict[str, Any]:
        try:
            url = f"https://graph.microsoft.com/v1.0/me/contacts/{contact_id}"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()

            contact_details = response.json()
            logger.info(f"Fetched contact details for ID: {contact_id}")
            return {"result": contact_details, "error": None}
        except requests.exceptions.RequestException as e:
            error_message = f"API request failed: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}
        except Exception as e:
            error_message = f"Error fetching contact details: {e}"
            logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}


def main(connection_id: str, contact_id: str) -> Dict[str, Any]:
    try:
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]

        contact_fetcher = OutlookContactFetcher()
        result = contact_fetcher.get_contact_by_id(
            access_token=access_token, contact_id=contact_id
        )
        return result
    except Exception as e:
        error_message = f"Error in main function: {e}"
        logger.error(error_message,extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}
