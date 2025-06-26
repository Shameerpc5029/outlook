#!/usr/bin/env python3
"""
Outlook MCP Server - A Model Context Protocol server for Microsoft Outlook integration.

This server provides tools for managing emails, contacts, calendars, and folders via Microsoft Graph API.
It implements the MCP stdio transport protocol for seamless integration with MCP clients.
"""

import asyncio
import json
import sys
from typing import Any, Dict, List

from mcp.server.stdio import stdio_server
from mcp.server import Server
from mcp.types import (
    TextContent,
    Tool,
)

# Import tool functions
from .tools.email import (
    send_email, create_draft_email, send_draft_email, get_draft_emails,
    delete_draft_email, update_draft_email,
)
from .tools.contacts import (
    create_contact, get_all_contacts, get_contact_details, update_contact, delete_contact,
)
from .tools.calendar import (
    get_all_calendars, get_calendar_details, create_calendar, update_calendar,
    delete_calendar, get_all_events, get_event_details, create_event, delete_event,
)
from .tools.folders import (
    get_all_folders, get_folder_details, create_folder, update_folder,
    delete_folder, get_many_folders,
)


class OutlookMCPServer:
    """MCP Server for Outlook integration using proper MCP patterns."""
    
    def __init__(self):
        """Initialize the Outlook MCP server with all tools."""
        self.server = Server("outlook-mcp")
        self._setup_tools()
    
    def _setup_tools(self):
        """Setup all available tools with their schemas."""
        
        # Email Tools
        @self.server.list_tools()
        async def list_tools() -> List[Tool]:
            """List all available tools."""
            return [
                # Email tools
                Tool(
                    name="send_email",
                    description="Send an email via Outlook with support for TO/CC/BCC recipients, HTML/text content, attachments, and importance levels",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string", "description": "Email subject line"},
                            "content": {"type": "string", "description": "Email body content (HTML or plain text)"},
                            "to_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of TO recipient email addresses"},
                            "cc_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of CC recipient email addresses (optional)"},
                            "bcc_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of BCC recipient email addresses (optional)"},
                            "content_type": {"type": "string", "enum": ["HTML", "Text"], "default": "HTML", "description": "Content type of email body"},
                            "save_to_sent": {"type": "boolean", "default": True, "description": "Whether to save email to sent items"},
                            "importance": {"type": "string", "enum": ["low", "normal", "high"], "default": "normal", "description": "Email importance level"}
                        },
                        "required": ["subject", "content", "to_recipients"]
                    }
                ),
                Tool(
                    name="create_draft_email",
                    description="Create a draft email that can be edited and sent later",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string", "description": "Email subject line"},
                            "content": {"type": "string", "description": "Email body content"},
                            "to_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of TO recipients"},
                            "cc_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of CC recipients (optional)"},
                            "bcc_recipients": {"type": "array", "items": {"type": "string"}, "description": "List of BCC recipients (optional)"},
                            "content_type": {"type": "string", "enum": ["HTML", "Text"], "default": "HTML"},
                            "importance": {"type": "string", "enum": ["low", "normal", "high"], "default": "normal"}
                        },
                        "required": ["subject", "content", "to_recipients"]
                    }
                ),
                Tool(
                    name="send_draft_email",
                    description="Send an existing draft email",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "draft_id": {"type": "string", "description": "ID of the draft email to send"}
                        },
                        "required": ["draft_id"]
                    }
                ),
                Tool(
                    name="get_draft_emails",
                    description="Retrieve all draft emails from the drafts folder",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                ),
                Tool(
                    name="update_draft_email",
                    description="Update an existing draft email",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "draft_id": {"type": "string", "description": "ID of the draft email to update"},
                            "subject": {"type": "string", "description": "Updated subject line"},
                            "content": {"type": "string", "description": "Updated content"},
                            "to_recipients": {"type": "array", "items": {"type": "string"}},
                            "cc_recipients": {"type": "array", "items": {"type": "string"}},
                            "bcc_recipients": {"type": "array", "items": {"type": "string"}},
                            "content_type": {"type": "string", "enum": ["HTML", "Text"]},
                            "importance": {"type": "string", "enum": ["low", "normal", "high"]}
                        },
                        "required": ["draft_id"]
                    }
                ),
                Tool(
                    name="delete_draft_email",
                    description="Delete a draft email",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "draft_id": {"type": "string", "description": "ID of the draft email to delete"}
                        },
                        "required": ["draft_id"]
                    }
                ),
                
                # Contact Tools
                Tool(
                    name="create_contact",
                    description="Create a new contact in Outlook with personal and business information",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "given_name": {"type": "string", "description": "First name of the contact"},
                            "surname": {"type": "string", "description": "Last name of the contact"},
                            "email_addresses": {"type": "string", "description": "Comma-separated email addresses"},
                            "business_phones": {"type": "string", "description": "Business phone numbers"},
                            "mobile_phone": {"type": "string", "description": "Mobile phone number"},
                            "job_title": {"type": "string", "description": "Job title"},
                            "company_name": {"type": "string", "description": "Company name"},
                            "department": {"type": "string", "description": "Department"},
                            "office_location": {"type": "string", "description": "Office location"}
                        },
                        "required": ["given_name"]
                    }
                ),
                Tool(
                    name="get_all_contacts",
                    description="Retrieve all contacts from Outlook",
                    inputSchema={"type": "object", "properties": {}}
                ),
                Tool(
                    name="get_contact_details",
                    description="Get detailed information about a specific contact",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "contact_id": {"type": "string", "description": "Unique identifier of the contact"}
                        },
                        "required": ["contact_id"]
                    }
                ),
                Tool(
                    name="update_contact",
                    description="Update an existing contact's information",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "contact_id": {"type": "string", "description": "Unique identifier of the contact"},
                            "given_name": {"type": "string"},
                            "surname": {"type": "string"},
                            "email_addresses": {"type": "string"},
                            "business_phones": {"type": "string"},
                            "mobile_phone": {"type": "string"},
                            "job_title": {"type": "string"},
                            "company_name": {"type": "string"},
                            "department": {"type": "string"},
                            "office_location": {"type": "string"}
                        },
                        "required": ["contact_id"]
                    }
                ),
                Tool(
                    name="delete_contact",
                    description="Delete a contact from Outlook",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "contact_id": {"type": "string", "description": "Unique identifier of the contact to delete"}
                        },
                        "required": ["contact_id"]
                    }
                ),
                
                # Calendar Tools
                Tool(
                    name="get_all_calendars",
                    description="Retrieve all calendars from Outlook",
                    inputSchema={"type": "object", "properties": {}}
                ),
                Tool(
                    name="get_calendar_details",
                    description="Get detailed information about a specific calendar",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "calendar_id": {"type": "string", "description": "Unique identifier of the calendar"}
                        },
                        "required": ["calendar_id"]
                    }
                ),
                Tool(
                    name="create_calendar",
                    description="Create a new calendar with specified name and color",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Name of the new calendar"},
                            "color": {"type": "string", "description": "Calendar color theme", "default": "auto"}
                        },
                        "required": ["name"]
                    }
                ),
                Tool(
                    name="update_calendar",
                    description="Update an existing calendar's properties",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "calendar_id": {"type": "string", "description": "Unique identifier of the calendar"},
                            "name": {"type": "string", "description": "Updated calendar name"},
                            "color": {"type": "string", "description": "Updated calendar color"}
                        },
                        "required": ["calendar_id"]
                    }
                ),
                Tool(
                    name="delete_calendar",
                    description="Delete a calendar from Outlook",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "calendar_id": {"type": "string", "description": "Unique identifier of the calendar to delete"}
                        },
                        "required": ["calendar_id"]
                    }
                ),
                Tool(
                    name="get_all_events",
                    description="Retrieve all events from a calendar or the default calendar",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "calendar_id": {"type": "string", "description": "Calendar ID (optional, uses default calendar if not specified)"}
                        }
                    }
                ),
                Tool(
                    name="get_event_details",
                    description="Get detailed information about a specific event",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "event_id": {"type": "string", "description": "Unique identifier of the event"}
                        },
                        "required": ["event_id"]
                    }
                ),
                Tool(
                    name="create_event",
                    description="Create a new calendar event with attendees and location",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string", "description": "Event title/subject"},
                            "start_datetime": {"type": "string", "description": "Start date and time in ISO format (e.g., 2024-01-15T10:00:00)"},
                            "end_datetime": {"type": "string", "description": "End date and time in ISO format"},
                            "start_timezone": {"type": "string", "default": "UTC", "description": "Start timezone"},
                            "end_timezone": {"type": "string", "default": "UTC", "description": "End timezone"},
                            "body_content": {"type": "string", "description": "Event description/body"},
                            "body_content_type": {"type": "string", "enum": ["HTML", "Text"], "default": "HTML"},
                            "location": {"type": "string", "description": "Event location"},
                            "attendees": {"type": "array", "items": {"type": "string"}, "description": "List of attendee email addresses"},
                            "calendar_id": {"type": "string", "description": "Calendar ID (optional, uses default calendar if not specified)"}
                        },
                        "required": ["subject", "start_datetime", "end_datetime"]
                    }
                ),
                Tool(
                    name="delete_event",
                    description="Delete an event from the calendar",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "event_id": {"type": "string", "description": "Unique identifier of the event to delete"}
                        },
                        "required": ["event_id"]
                    }
                ),
                
                # Folder Tools
                Tool(
                    name="get_all_folders",
                    description="Retrieve all mail folders from Outlook",
                    inputSchema={"type": "object", "properties": {}}
                ),
                Tool(
                    name="get_folder_details",
                    description="Get detailed information about a specific mail folder",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_id": {"type": "string", "description": "Unique identifier of the folder"}
                        },
                        "required": ["folder_id"]
                    }
                ),
                Tool(
                    name="create_folder",
                    description="Create a new mail folder, optionally nested under a parent folder",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "display_name": {"type": "string", "description": "Display name for the new folder"},
                            "parent_folder_id": {"type": "string", "description": "Parent folder ID (optional, creates in root if not specified)"}
                        },
                        "required": ["display_name"]
                    }
                ),
                Tool(
                    name="update_folder",
                    description="Update a folder's display name",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_id": {"type": "string", "description": "Unique identifier of the folder"},
                            "display_name": {"type": "string", "description": "New display name for the folder"}
                        },
                        "required": ["folder_id", "display_name"]
                    }
                ),
                Tool(
                    name="delete_folder",
                    description="Delete a mail folder from Outlook",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_id": {"type": "string", "description": "Unique identifier of the folder to delete"}
                        },
                        "required": ["folder_id"]
                    }
                ),
                Tool(
                    name="get_many_folders",
                    description="Get detailed information for multiple folders in a single request",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_ids": {"type": "array", "items": {"type": "string"}, "description": "List of folder IDs to retrieve"}
                        },
                        "required": ["folder_ids"]
                    }
                )
            ]
        
        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
            """Execute a tool with the given arguments."""
            try:
                # Route the tool call to the appropriate function
                if name == "send_email":
                    result = send_email(**arguments)
                elif name == "create_draft_email":
                    result = create_draft_email(**arguments)
                elif name == "send_draft_email":
                    result = send_draft_email(**arguments)
                elif name == "get_draft_emails":
                    result = get_draft_emails(**arguments)
                elif name == "update_draft_email":
                    result = update_draft_email(**arguments)
                elif name == "delete_draft_email":
                    result = delete_draft_email(**arguments)
                elif name == "create_contact":
                    result = create_contact(**arguments)
                elif name == "get_all_contacts":
                    result = get_all_contacts(**arguments)
                elif name == "get_contact_details":
                    result = get_contact_details(**arguments)
                elif name == "update_contact":
                    result = update_contact(**arguments)
                elif name == "delete_contact":
                    result = delete_contact(**arguments)
                elif name == "get_all_calendars":
                    result = get_all_calendars(**arguments)
                elif name == "get_calendar_details":
                    result = get_calendar_details(**arguments)
                elif name == "create_calendar":
                    result = create_calendar(**arguments)
                elif name == "update_calendar":
                    result = update_calendar(**arguments)
                elif name == "delete_calendar":
                    result = delete_calendar(**arguments)
                elif name == "get_all_events":
                    result = get_all_events(**arguments)
                elif name == "get_event_details":
                    result = get_event_details(**arguments)
                elif name == "create_event":
                    result = create_event(**arguments)
                elif name == "delete_event":
                    result = delete_event(**arguments)
                elif name == "get_all_folders":
                    result = get_all_folders(**arguments)
                elif name == "get_folder_details":
                    result = get_folder_details(**arguments)
                elif name == "create_folder":
                    result = create_folder(**arguments)
                elif name == "update_folder":
                    result = update_folder(**arguments)
                elif name == "delete_folder":
                    result = delete_folder(**arguments)
                elif name == "get_many_folders":
                    result = get_many_folders(**arguments)
                else:
                    raise ValueError(f"Unknown tool: {name}")
                
                # Return the result as TextContent
                return [TextContent(type="text", text=json.dumps(result, indent=2, default=str))]
                
            except Exception as e:
                # Return error information
                error_result = {
                    "error": str(e),
                    "tool": name,
                    "arguments": arguments
                }
                return [TextContent(type="text", text=json.dumps(error_result, indent=2))]
    
    async def run(self):
        """Run the MCP server using stdio transport."""
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream, 
                write_stream, 
                self.server.create_initialization_options()
            )


def main():
    """Main entry point for the MCP server."""
    import asyncio
    import sys
    
    # Check if we're being run directly or as a module
    if len(sys.argv) > 1 and sys.argv[1] == "--help":
        print("Outlook MCP Server")
        print("A Model Context Protocol server for Microsoft Outlook integration.")
        print("")
        print("Usage:")
        print("  python -m outlook_mcp.server")
        print("  or")
        print("  outlook-mcp")
        print("")
        print("This server provides 26 tools for managing:")
        print("  • Emails (send, draft, update)")
        print("  • Contacts (create, read, update, delete)")
        print("  • Calendars and Events (full CRUD operations)")
        print("  • Mail Folders (organize and manage)")
        print("")
        print("The server communicates via stdin/stdout using the MCP protocol.")
        return
    
    # Create and run the server
    server = OutlookMCPServer()
    try:
        asyncio.run(server.run())
    except KeyboardInterrupt:
        print("\\nServer shutting down...", file=sys.stderr)
    except Exception as e:
        print(f"Server error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
