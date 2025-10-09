[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/ryaker-outlook-mcp-badge.png)](https://mseep.ai/app/ryaker-outlook-mcp)

# Modular Outlook MCP Server

This is a modular implementation of the Outlook MCP (Model Context Protocol) server that connects Claude with Microsoft Outlook through the Microsoft Graph API.
Certified by MCPHub https://mcphub.com/mcp-servers/ryaker/outlook-mcp

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
- **Calendar Management**: List, create, accept, decline, and delete calendar events
- **Modular Structure**: Clean separation of concerns for better maintainability
- **OData Filter Handling**: Proper escaping and formatting of OData queries
- **Test Mode**: Simulated responses for testing without real API calls

## Quick Start

1. **Install dependencies**: `npm install`
2. **Azure setup**: Register app in Azure Portal (see detailed steps below)
3. **Configure environment**: Copy `.env.example` to `.env` and add your Azure credentials
4. **Configure Claude**: Update your Claude Desktop config with the server path
5. **Start auth server**: `npm run auth-server` 
6. **Authenticate**: Use the authenticate tool in Claude to get the OAuth URL
7. **Start using**: Access your Outlook data through Claude!

## Installation

### Prerequisites
- Node.js 14.0.0 or higher
- npm or yarn package manager
- Azure account for app registration

### Install Dependencies

```bash
npm install
```

This will install the required dependencies including:
- `@modelcontextprotocol/sdk` - MCP protocol implementation
- `dotenv` - Environment variable management

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
7. **⚠️ IMPORTANT**: Copy the secret **VALUE** (not the Secret ID) and save it for the next step

## Configuration

### 1. Environment Variables

Create a `.env` file in the project root by copying the example:

```bash
cp .env.example .env
```

Edit `.env` and add your Azure credentials:

```bash
# Get these values from Azure Portal > App Registrations > Your App
MS_CLIENT_ID=your-application-client-id-here
MS_CLIENT_SECRET=your-client-secret-VALUE-here
USE_TEST_MODE=false
```

**Important Notes:**
- Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in the `.env` file
- For Claude Desktop config, you'll use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- Always use the client secret **VALUE**, never the Secret ID

### 2. Claude Desktop Configuration

Copy the configuration from `claude-config-sample.json` to your Claude Desktop config file and update the paths and credentials:

```json
{
  "mcpServers": {
    "outlook-assistant": {
      "command": "node",
      "args": [
        "/absolute/path/to/outlook-mcp/index.js"
      ],
      "env": {
        "USE_TEST_MODE": "false",
        "OUTLOOK_CLIENT_ID": "your-client-id-here",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret-here"
      }
    }
  }
}
```

### 3. Advanced Configuration (Optional)

To configure server behavior, you can edit `config.js` to change:

- Server name and version
- Test mode settings
- Authentication parameters
- Email field selections
- API endpoints

## Usage with Claude Desktop

1. **Configure Claude Desktop**: Add the server configuration (see Configuration section above)
2. **Restart Claude Desktop**: Close and reopen Claude Desktop to load the new MCP server
3. **Start Authentication Server**: Open a terminal and run `npm run auth-server`
4. **Authenticate**: In Claude Desktop, use the `authenticate` tool to get an OAuth URL
5. **Complete OAuth Flow**: Visit the URL in your browser and sign in with Microsoft
6. **Start Using**: Once authenticated, you can use all the Outlook tools in Claude!

## Running Standalone

You can test the server using:

```bash
./test-modular-server.sh
```

This will use the MCP Inspector to directly connect to the server and let you test the available tools.

## Authentication Flow

The authentication process requires two steps:

### Step 1: Start the Authentication Server
```bash
npm run auth-server
```
This starts a local server on port 3333 that handles the OAuth callback from Microsoft.

**⚠️ Important**: The auth server MUST be running before you try to authenticate. The authentication URL will not work if the server isn't running.

### Step 2: Authenticate with Microsoft
1. In Claude Desktop, use the `authenticate` tool
2. Claude will provide a URL like: `http://localhost:3333/auth?client_id=your-client-id`
3. Visit this URL in your browser
4. Sign in with your Microsoft account
5. Grant the requested permissions
6. You'll be redirected back to a success page
7. Tokens are automatically stored in `~/.outlook-mcp-tokens.json`

The authentication server can be stopped after successful authentication (tokens are saved). However, you'll need to restart it if you need to re-authenticate.

## Troubleshooting

### Common Installation Issues

#### "Cannot find module '@modelcontextprotocol/sdk/server/index.js'"
**Solution**: Install dependencies first:
```bash
npm install
```

#### "Error: listen EADDRINUSE: address already in use :::3333"
**Solution**: Port 3333 is already in use. Kill the existing process:
```bash
npx kill-port 3333
```
Then restart the auth server: `npm run auth-server`

### Authentication Issues

#### "Invalid client secret provided" (Error AADSTS7000215)
**Root Cause**: You're using the Secret ID instead of the Secret Value.

**Solution**:
1. Go to Azure Portal > App Registrations > Your App > Certificates & secrets
2. Copy the **Value** column (not the Secret ID column)
3. Update both:
   - `.env` file: `MS_CLIENT_SECRET=actual-secret-value`
   - Claude Desktop config: `OUTLOOK_CLIENT_SECRET=actual-secret-value`
4. Restart the auth server: `npm run auth-server`

#### Authentication URL doesn't work / "This site can't be reached"
**Root Cause**: Authentication server isn't running.

**Solution**:
1. Start the auth server first: `npm run auth-server`
2. Wait for "Authentication server running at http://localhost:3333"
3. Then try the authentication URL in Claude

#### "Authentication required" after successful setup
**Root Cause**: Token may have expired or been corrupted.

**Solutions**:
1. Check if token file exists: `~/.outlook-mcp-tokens.json`
2. If corrupted, delete the file and re-authenticate
3. Restart the auth server and authenticate again

### Configuration Issues

#### Server doesn't start in Claude Desktop
**Solutions**:
1. Check the absolute path in your Claude Desktop config
2. Ensure `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET` are set in Claude config
3. Restart Claude Desktop after config changes

#### Environment variables not loading
**Solutions**:
1. Ensure `.env` file exists in the project root
2. Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET` in `.env`
3. Don't add quotes around values in `.env` file

### API and Runtime Issues

- **OData Filter Errors**: Check server logs for escape sequence issues
- **API Call Failures**: Look for detailed error messages in the response
- **Token Refresh Issues**: Delete `~/.outlook-mcp-tokens.json` and re-authenticate

### Getting Help

If you're still having issues:
1. Check the console output from `npm run auth-server` for detailed error messages
2. Verify your Azure app registration settings match the documentation
3. Ensure you have the required Microsoft Graph API permissions

## Extending the Server

To add more functionality:

1. Create new module directories (e.g., `calendar/`)
2. Implement tool handlers in separate files
3. Export tool definitions from module index files
4. Import and add tools to `TOOLS` array in `index.js`
