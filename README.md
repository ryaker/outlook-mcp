# Modular Outlook MCP Server

This is a modular implementation of the Outlook MCP (Model Context Protocol) server that connects Claude with Microsoft Outlook through the Microsoft Graph API.

## Directory Structure

```
/modular/
├── index.js                 # Main entry point
├── config.js                # Configuration settings
├── auth/                    # Authentication modules
│   ├── index.js             # Authentication exports
│   ├── token-manager.js     # Token storage and refresh
│   └── tools.js             # Auth-related tools
├── calendar/                # Calendar functionality
│   ├── index.js             # Calendar exports
│   ├── list.js              # List events
│   ├── create.js            # Create event
│   ├── delete.js            # Delete event
│   ├── cancel.js            # Cancel
│   ├── accept.js            # Accept event
│   ├── tentative.js         # Tentatively accept event
│   ├── decline.js           # Decline event
├── email/                   # Email functionality
│   ├── index.js             # Email exports
│   ├── list.js              # List emails
│   ├── search.js            # Search emails
│   ├── read.js              # Read email
│   └── send.js              # Send email
└── utils/                   # Utility functions
    ├── graph-api.js         # Microsoft Graph API helper
    ├── odata-helpers.js     # OData query building
    └── mock-data.js         # Test mode data
```

## Features

- **Authentication**: OAuth 2.0 authentication with Microsoft Graph API
- **Email Management**: List, search, read, and send emails
- **Modular Structure**: Clean separation of concerns for better maintainability
- **OData Filter Handling**: Proper escaping and formatting of OData queries
- **Test Mode**: Simulated responses for testing without real API calls

## Azure App Registration & Configuration

To use this MCP server you need to first register and configure an app in Azure Portal. The following steps will take you through the process of registering a new app, configuring its permissions, and generating a client secret.

### App Registration

1. Open [Azure Portal](https://portal.azure.com/) in your browser
2. Sign in with a Microsoft Work or Personal account
3. Search for or cilck on "App registrations"
4. Click on "New registration"
5. Enter a name for the app, for example "Outlook MCP Server"
6. Select the "Accounts in any organizational directory and personal Microsoft accounts" option
7. In the "Redirect URI" section, select "Web" from the dropdown and enter "http://localhost:3333/auth/callback" in the textbox
8. Click on "Register"
9. From the Overview section of the app settings page, copy the "Application (client) ID" and enter it as the MS_CLIENT_ID in the .env file as well as the OUTLOOK_CLIENT_ID in the claude-config-sample.json file

### App Permissions

1. From the app settings page in Azure Portal select the "API permissions" option under the Manage section
2. Click on "Add a permission"
3. Click on "Microsoft Graph"
4. Select "Delegated permissions"
5. Search for the following permissions and slect the checkbox next to each one
    - offline_access
    - User.Read
    - Mail.Read
    - Mail.Send
    - Calendars.Read
    - Calendars.ReadWrite
    - Contacts.Read
6. Click on "Add permissions"

### Client Secret

1. From the app settings page in Azure Portal select the "Certificates & secrets" option under the Manage section
2. Switch to the "Client secrets" tab
3. Click on "New client secret"
4. Enter a description, for example "Client Secret"
5. Select the longest possible expiration time
6. Click on "Add"
7. Copy the secret value and enter it as the MS_CLIENT_SECRET in the .env file as well as the OUTLOOK_CLIENT_SECRET in the claude-config-sample.json file

## Configuration

To configure the server, edit the `config.js` file to change:

- Server name and version
- Test mode settings
- Authentication parameters
- Email field selections
- API endpoints

## Usage with Claude Desktop

1. Copy the sample configuration from `claude-config-sample.json` to your Claude Desktop configuration
2. Restart Claude Desktop
3. Authenticate with Microsoft using the `authenticate` tool
4. Use the email tools to manage your Outlook account

## Running Standalone

You can test the server using:

```bash
./test-modular-server.sh
```

This will use the MCP Inspector to directly connect to the server and let you test the available tools.

## SSE Mode

To run the server in SSE mode, use the following command:

```
node sse-server.js
```

This will start the server in SSE mode. The SSE server also integrates with the auth endpoint and handles
auth callbacks.

## Authentication Flow

1. Start a local authentication server on port 3333 (using `outlook-auth-server.js`)
2. Use the `authenticate` tool to get an authentication URL
3. Complete the authentication in your browser
4. Tokens are stored in `~/.outlook-mcp-tokens.json`

## Troubleshooting

- **Authentication Issues**: Check the token file and authentication server logs
- **OData Filter Errors**: Look for escape sequences in the server logs
- **API Call Failures**: Check for detailed error messages in the response

## Extending the Server

To add more functionality:

1. Create new module directories (e.g., `calendar/`)
2. Implement tool handlers in separate files
3. Export tool definitions from module index files
4. Import and add tools to `TOOLS` array in `index.js`
