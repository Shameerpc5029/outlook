from typing import List, Dict, Any
import requests
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger
import os


class OutlookFoldersFetcher:
    @staticmethod
    def get_folders_by_ids(access_token: str, folder_ids: List[str]) -> Dict[str, Any]:
        try:
            url = "https://graph.microsoft.com/v1.0/$batch"
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
            }

            # Build batch request payload
            requests_payload = [
                {
                    "id": f"request_{i}",
                    "method": "GET",
                    "url": f"/me/mailFolders/{folder_id}",
                }
                for i, folder_id in enumerate(folder_ids)
            ]
            batch_payload = {"requests": requests_payload}

            logger.info(
                "Sending batch request to fetch folder details.",
                extra={"url": url, "batch_payload": batch_payload},
            )

            response = requests.post(
                url, headers=headers, json=batch_payload, timeout=10
            )
            response.raise_for_status()

            batch_response = response.json()

            folder_details = [
                {"id": item["id"], "status": item["status"], "body": item.get("body")}
                for item in batch_response.get("responses", [])
            ]

            logger.info(
                "Successfully fetched folder details.",
                extra={"folder_ids": folder_ids, "folder_details": folder_details},
            )

            return {"result": folder_details, "error": None}
        except requests.exceptions.RequestException as e:
            logger.error(
                "API request failed.",
                extra={
                    "folder_ids": folder_ids,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}
        except Exception as e:
            logger.error(
                "Error fetching folder details.",
                extra={
                    "folder_ids": folder_ids,
                    "error": str(e),
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": str(e)}


def main(connection_id: str, folder_ids: List[str]) -> dict[str, Any]:
    try:
        logger.info(
            "Fetching credentials for the provided connection ID.",
            extra={"connection_id": connection_id},
        )

        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )

        if "credentials" not in credentials:
            error_message = (
                "Missing 'credentials' in the response from get_connection_credentials."
            )
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "credentials_response": credentials,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        access_token = credentials["credentials"].get("access_token")
        if not access_token:
            error_message = "Access token is missing in the credentials."
            logger.error(
                error_message,
                extra={
                    "connection_id": connection_id,
                    "credentials": credentials,
                    "path": os.getenv("WM_JOB_PATH"),
                },
            )
            return {"result": None, "error": error_message}

        logger.info(
            "Starting process to fetch multiple folder details.",
            extra={"folder_ids": folder_ids},
        )

        folder_fetcher = OutlookFoldersFetcher()
        result = folder_fetcher.get_folders_by_ids(
            access_token=access_token, folder_ids=folder_ids
        )
        return result
    except Exception as e:
        logger.error(
            "Error in main function.",
            extra={"error": str(e), "path": os.getenv("WM_JOB_PATH")},
        )
        return {"result": None, "error": str(e)}
