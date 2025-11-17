# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `npm install` - **ALWAYS run first** to install dependencies
- `npm start` - Start the MCP server
- `npm start:clean` - Start services with automatic port cleanup (recommended)
- `npm run auth-server` - Start the OAuth authentication server on port 3333 (**required for authentication**)
- `npm run cleanup` - Clean up ports 3333, 6274, 6277 and kill stale processes
- `npm run test-mode` - Start the server in test mode with mock data
- `npm run inspect` - Use MCP Inspector to test the server interactively
- `npm test` - Run Jest tests
- `./test-modular-server.sh` - Test the server using MCP Inspector
- `./test-direct.sh` - Direct testing script
- `npx kill-port 3333` - Kill process using port 3333 if auth server won't start

### Port Conflict Prevention

The system includes automatic port conflict prevention:
- **Cleanup Script**: `scripts/cleanup-ports.sh` - Gracefully shuts down old processes before starting new ones
- **Startup Script**: `scripts/start-services.sh` - Runs cleanup, then starts both auth server and MCP server
- **Graceful Shutdown**: The MCP server now handles SIGTERM and SIGINT signals properly
- **launchd Configuration**: Updated to use the cleanup script on startup and throttle rapid restarts

**To resolve port conflicts manually:**
```bash
npm run cleanup
```

## Architecture Overview

This is a modular MCP (Model Context Protocol) server that provides Claude with access to Microsoft Outlook via the Microsoft Graph API. The architecture is organized into functional modules:

### Core Structure
- `index.js` - Main entry point that combines all module tools and handles MCP protocol
- `config.js` - Centralized configuration including API endpoints, field selections, and authentication settings
- `outlook-auth-server.js` - Standalone OAuth server for authentication flow

### Modules
Each module exports tools and handlers:
- `auth/` - OAuth 2.0 authentication with token management
- `calendar/` - Calendar operations (list, create, accept, decline, delete events)
- `email/` - Email management (list, search, read, send, mark as read)
- `folder/` - Folder operations (list, create, move)
- `rules/` - Email rules management
- `utils/` - Shared utilities including Graph API client and OData helpers

### Key Components
- **Token Management**: Tokens stored in `~/.outlook-mcp-tokens.json`
- **Graph API Client**: `utils/graph-api.js` handles all Microsoft Graph API calls with proper OData encoding
- **Test Mode**: Mock data responses when `USE_TEST_MODE=true`
- **Modular Tools**: Each module exports tools array that gets combined in main server

## Authentication Flow

1. Azure app registration required with specific permissions (Mail.Read, Mail.Send, Calendars.ReadWrite, etc.)
2. Start auth server: `npm run auth-server` 
3. Use authenticate tool to get OAuth URL
4. Complete browser authentication
5. Tokens automatically stored and refreshed

## Configuration Requirements

### Environment Variables
- **For .env file**: Use `MS_CLIENT_ID` and `MS_CLIENT_SECRET`
- **For Claude Desktop config**: Use `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET`
- **Important**: Always use the client secret VALUE from Azure, not the Secret ID
- Copy `.env.example` to `.env` and populate with real Azure credentials
- Default timezone is "Central European Standard Time"
- Default page size is 25, max results 50

### Common Setup Issues
1. **Missing dependencies**: Always run `npm install` first
2. **Wrong secret**: Use Azure secret VALUE, not ID (AADSTS7000215 error)
3. **Auth server not running**: Start `npm run auth-server` before authenticating
4. **Port conflicts**:
   - Use `npm run cleanup` to automatically clean up ports
   - Or use `npx kill-port 3333` to kill specific ports
   - Multiple processes can accumulate - the cleanup script handles this gracefully
5. **Connection errors in MCP Inspector**: Run `npm run cleanup` then restart the inspector

## Test Mode

Set `USE_TEST_MODE=true` to use mock data instead of real API calls. Mock responses are defined in `utils/mock-data.js`.

## OData Query Handling

The Graph API client properly handles OData filters with URI encoding. Filters are processed separately from other query parameters to ensure correct escaping of special characters.

## Error Handling

- Authentication failures return "UNAUTHORIZED" error
- Graph API errors include status codes and response details
- Token expiration triggers re-authentication flow
- Empty API responses are handled gracefully (returns '{}' if empty)