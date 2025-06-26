from typing import Dict, Any, List
import requests
import os
from f.common.nango.connections import get_connection_credentials
from f.common.logfire.logger import logger

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
    def send_emails(
        access_token: str, emails_data: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        try:
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
                        logger.info(f"Email sent successfully to {', '.join(email_data.get('to', []))}")
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
                    logger.error(f"Error sending individual email: {e}", 
                               extra={"path": os.getenv("WM_JOB_PATH")})
                    results.append({
                        "status": "failed",
                        "recipients": email_data.get("to", []),
                        "error": str(e)
                    })
            
            return {"result": results, "error": None}
            
        except Exception as e:
            error_message = f"Error in batch email sending: {e}"
            logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
            return {"result": None, "error": error_message}

def main(
    connection_id: str,
    subject: str,
    content: str,
    to_recipients: List[str],
    cc_recipients: List[str] = None,
    bcc_recipients: List[str] = None,
    content_type: str = "HTML",
    save_to_sent: bool = True,
    importance: str = None,
    attachments: List[Dict[str, Any]] = None,
    custom_headers: List[Dict[str, str]] = None,
    flag: Dict[str, Any] = None
) -> Dict[str, Any]:
    """
    Windmill script to send emails via Outlook with all available options
    
    Args:
        connection_id: The Outlook connection ID from Nango
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
        credentials = get_connection_credentials(
            id=connection_id, providerConfigKey="outlook"
        )
        access_token = credentials["credentials"]["access_token"]
        email_sender = OutlookEmailSender()
        return email_sender.send_emails(
            access_token=access_token, 
            emails_data=[email_data]
        )
    except Exception as e:
        error_message = f"Error in email sending script: {e}"
        logger.error(error_message, extra={"path": os.getenv("WM_JOB_PATH")})
        return {"result": None, "error": error_message}


# This is sample email data with all options 

#      sample_email_data = {
#     "connection_id": "nango_connection_id",
#     "subject": "mail subject",
#     "content": """
#         <div style='font-family: Arial, sans-serif;'>
#             <h1>Q4 2023 Performance Review</h1>
#             <p>Dear Team,</p>
#             <p>Please find attached our Q4 2023 performance review documents.</p>
#             <h2>Key Highlights:</h2>
#             <ul>
#                 <li>Revenue growth: 15%</li>
#                 <li>Customer satisfaction: 94%</li>
#                 <li>New product launches: 3</li>
#             </ul>
#             <p>Best regards,<br>Management Team</p>
#         </div>
#     """,
#     "to_recipients": [
#         "executive.team@example.com",
#         "department.heads@example.com"
#     ],
#     "cc_recipients": [
#         "board.members@example.com",
#         "stakeholders@example.com"
#     ],
#     "bcc_recipients": [
#         "records@example.com",
#         "audit@example.com"
#     ],
#     "content_type": "HTML",
#     "save_to_sent": True,
#     "importance": "high",
#     "attachments": [
#         {
#             "name": "Q4_2023_Review.pdf",
#             "contentType": "application/pdf",
#             "contentBytes": base64 
#         },
#         {
#             "name": "Financial_Metrics.xlsx",
#             "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             "contentBytes": base64
#         }
#     ],
#     "custom_headers": [
#         {
#             "name": "X-Report-Type",
#             "value": "Quarterly Review"
#         },
#         {
#             "name": "X-Department",
#             "value": "Executive Management"
#         },
#         {
#             "name": "X-Classification",
#             "value": "Internal-Only"
#         }
#     ],
#     "flag": {
#         "flagStatus": "flagged",
#         "dueDateTime": {
#             "dateTime": datetime.now(timezone.utc).isoformat(),
#             "timeZone": "UTC"
#         }
#     }
# }