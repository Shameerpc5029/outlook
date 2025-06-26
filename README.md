# Outlook MCP Server

A **Model Context Protocol (MCP)** server for Microsoft Outlook integration that provides comprehensive email, contact, calendar, and folder management capabilities through the Microsoft Graph API.

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- Microsoft Graph API access via [Nango](https://nango.dev)
- Valid Outlook/Microsoft 365 account

### Installation

1. **Clone and setup:**
   ```bash
   git clone <repository-url>
   cd outlook-mcp
   pip install -e .
   ```

2. **Configure environment variables:**
   Create a `.env` file or set these environment variables:
   ```bash
   NANGO_CONNECTION_ID=your_connection_id
   NANGO_INTEGRATION_ID=your_integration_id
   NANGO_BASE_URL=your_nango_base_url
   NANGO_SECRET_KEY=your_secret_key
   ```

3. **Test the server:**
   ```bash
   python outlook_mcp_server.py --help
   ```

## 🔧 MCP Client Configuration

### Claude Desktop Configuration

Add this to your Claude Desktop MCP configuration:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "uvx",
      "args": [
        "git+https://github.com/Shameerpc5029/outlook.git"
      ],
      "env": {
        "NANGO_CONNECTION_ID": "your_connection_id",
        "NANGO_INTEGRATION_ID": "your_integration_id",
        "NANGO_BASE_URL": "your_nango_base_url",
        "NANGO_SECRET_KEY": "your_secret_key"
      }
    }
  }
}
```

### Other MCP Clients

For other MCP clients, use:
- **Command:** `python`
- **Args:** `["/path/to/outlook_mcp_server.py"]`
- **Transport:** stdio
- **Environment:** Set the required Nango variables

## 📧 Available Tools (26 Total)

### Email Management (6 tools)
- **`send_email`** - Send emails with TO/CC/BCC, HTML/text content, attachments
- **`create_draft_email`** - Create draft emails for later editing
- **`send_draft_email`** - Send existing draft emails
- **`get_draft_emails`** - Retrieve all draft emails
- **`update_draft_email`** - Modify existing draft emails
- **`delete_draft_email`** - Remove draft emails

### Contact Management (5 tools)
- **`create_contact`** - Add new contacts with full details
- **`get_all_contacts`** - Retrieve all contacts
- **`get_contact_details`** - Get specific contact information
- **`update_contact`** - Modify existing contact details
- **`delete_contact`** - Remove contacts

### Calendar Management (9 tools)
- **`get_all_calendars`** - List all calendars
- **`get_calendar_details`** - Get specific calendar information
- **`create_calendar`** - Create new calendars with custom colors
- **`update_calendar`** - Modify calendar properties
- **`delete_calendar`** - Remove calendars
- **`get_all_events`** - Retrieve events from calendars
- **`get_event_details`** - Get specific event information
- **`create_event`** - Schedule new events with attendees
- **`delete_event`** - Remove calendar events

### Folder Management (6 tools)
- **`get_all_folders`** - List all mail folders
- **`get_folder_details`** - Get specific folder information
- **`create_folder`** - Create new mail folders (with nesting)
- **`update_folder`** - Rename folders
- **`delete_folder`** - Remove folders
- **`get_many_folders`** - Batch retrieve multiple folders

## 💡 Usage Examples

### Send an Email
```json
{
  "tool": "send_email",
  "arguments": {
    "subject": "Project Update",
    "content": "<h1>Hello!</h1><p>Here's the latest update...</p>",
    "to_recipients": ["colleague@company.com"],
    "cc_recipients": ["manager@company.com"],
    "importance": "high"
  }
}
```

### Create a Contact
```json
{
  "tool": "create_contact",
  "arguments": {
    "given_name": "John",
    "surname": "Doe",
    "email_addresses": "john.doe@company.com,john@personal.com",
    "company_name": "Acme Corp",
    "job_title": "Software Engineer"
  }
}
```

### Schedule a Meeting
```json
{
  "tool": "create_event",
  "arguments": {
    "subject": "Weekly Team Meeting",
    "start_datetime": "2024-01-15T10:00:00",
    "end_datetime": "2024-01-15T11:00:00",
    "attendees": ["team@company.com"],
    "location": "Conference Room A",
    "body_content": "Discussing project milestones"
  }
}
```

## 🏗️ Project Structure

```
outlook-mcp/
├── src/outlook_mcp/
│   ├── __init__.py
│   ├── server.py              # Main MCP server implementation
│   ├── connection.py          # Nango API connection handling
│   └── tools/
│       ├── __init__.py
│       ├── email.py           # Email management tools
│       ├── contacts.py        # Contact management tools
│       ├── calendar.py        # Calendar and event tools
│       └── folders.py         # Folder management tools
├── outlook_mcp_server.py      # Standalone server entry point
├── main.py                    # Alternative entry point
├── mcp_config.json           # Example MCP configuration
├── pyproject.toml            # Package configuration
├── README.md                 # This file
└── .env                      # Environment variables (create this)
```

## 🔒 Security & Authentication

This server uses **Nango** for secure Microsoft Graph API authentication:

1. **No direct credential storage** - All auth handled by Nango
2. **Token management** - Automatic token refresh and management
3. **Secure communication** - HTTPS-only API communication
4. **Environment-based config** - Sensitive data in environment variables

## 🐛 Troubleshooting

### Common Issues

1. **"Missing environment variables"**
   - Ensure all 4 Nango variables are set
   - Check `.env` file exists and is properly formatted

2. **"Connection failed"**
   - Verify Nango integration is active
   - Check internet connectivity
   - Validate Nango credentials

3. **"Tool execution failed"**
   - Check Microsoft Graph API permissions
   - Verify Outlook account has necessary access
   - Review error messages in server logs

### Debug Mode

Run with verbose output:
```bash
python outlook_mcp_server.py 2>debug.log
```

### Testing Connection

```bash
python -c "from outlook_mcp.connection import get_access_token; print('✅ Connection successful!' if get_access_token() else '❌ Connection failed')"
```

## 🧪 Development & Testing

### Local Development

1. **Install in development mode:**
   ```bash
   pip install -e .
   ```

2. **Run tests:**
   ```bash
   python -m pytest tests/
   ```

3. **Check tool functionality:**
   ```bash
   python -c "from outlook_mcp.tools.email import get_draft_emails; print(get_draft_emails())"
   ```

### Creating Custom Tools

1. Add your tool function to the appropriate module in `tools/`
2. Update the server's tool list in `server.py`
3. Add proper input schema validation
4. Test with MCP client

## 📋 Requirements

- **Python 3.8+**
- **mcp >= 1.0.0** - Model Context Protocol library
- **requests >= 2.32.4** - HTTP client
- **python-dotenv >= 1.1.1** - Environment variable management

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-tool`
3. Make your changes with tests
4. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🔗 Related Links

- [Model Context Protocol Specification](https://spec.modelcontextprotocol.io/)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Nango Integration Platform](https://nango.dev)
- [Claude Desktop MCP Configuration](https://docs.anthropic.com/claude/desktop)

## 📞 Support

For issues and questions:
1. Check the troubleshooting section above
2. Review the GitHub issues
3. Create a new issue with detailed information

---

**Built with ❤️ for seamless Outlook integration via MCP**
