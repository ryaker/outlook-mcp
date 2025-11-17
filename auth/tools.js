/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const TokenStorage = require('./token-storage');

// Initialize TokenStorage for multi-account support
const tokenStorage = new TokenStorage(config.AUTH_CONFIG);

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `📧 MODULAR Outlook Assistant MCP Server v${config.SERVER_VERSION} 📧\n\nProvides access to Microsoft Outlook email, calendar, and contacts through Microsoft Graph API.\nImplemented with a modular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    const testTokens = {
      access_token: "test_access_token_" + Date.now(),
      refresh_token: "test_refresh_token_" + Date.now(),
      expires_at: Date.now() + (3600 * 1000)
    };

    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  // Check if already authenticated (unless forcing re-auth)
  if (!force) {
    try {
      const token = await tokenStorage.getValidAccessToken();
      const activeAccount = await tokenStorage.getActiveAccount();
      if (token) {
        return {
          content: [{
            type: "text",
            text: `✓ Already authenticated as: ${activeAccount}`
          }]
        };
      }
    } catch (error) {
      console.error('[AUTHENTICATE] Error checking existing tokens:', error);
    }
  }

  // For real authentication, generate an auth URL and instruct the user to visit it
  const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;

  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');

  try {
    const token = await tokenStorage.getValidAccessToken();
    const activeAccount = await tokenStorage.getActiveAccount();
    const accounts = await tokenStorage.getAllAccounts();

    if (!token) {
      console.error('[CHECK-AUTH-STATUS] No valid access token found');
      return {
        content: [{ type: "text", text: "Not authenticated" }]
      };
    }

    console.error('[CHECK-AUTH-STATUS] Access token present');
    console.error(`[CHECK-AUTH-STATUS] Active account: ${activeAccount}`);
    console.error(`[CHECK-AUTH-STATUS] Available accounts: ${accounts.join(', ')}`);

    return {
      content: [{
        type: "text",
        text: `Authenticated and ready\n\nActive Account: ${activeAccount}\nAvailable Accounts: ${accounts.join(', ')}`
      }]
    };
  } catch (error) {
    console.error('[CHECK-AUTH-STATUS] Error:', error);
    return {
      content: [{ type: "text", text: `Error checking authentication: ${error.message}` }]
    };
  }
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this Outlook Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
