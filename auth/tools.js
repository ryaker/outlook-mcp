/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');

// Lazy-loaded to avoid circular dependency — tokenStorage is created in index.js
let _tokenStorage = null;
function getTokenStorage() {
  if (!_tokenStorage) {
    _tokenStorage = require('./index').tokenStorage;
  }
  return _tokenStorage;
}

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nAuth flow: ${config.AUTH_FLOW}`
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
    tokenManager.createTestTokens();
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  // Device code flow
  if (config.AUTH_FLOW === 'device_code') {
    try {
      const ts = getTokenStorage();

      // Check if already authenticated (unless forcing)
      if (!force) {
        // Don't start a new flow if polling is already in progress
        if (ts._pendingDeviceCodePoll) {
          return {
            content: [{
              type: "text",
              text: 'Authentication is already in progress. Complete the sign-in in your browser, or use check-auth-status to verify.'
            }]
          };
        }
        const existing = await ts.getValidAccessToken();
        if (existing) {
          return {
            content: [{
              type: "text",
              text: 'Already authenticated. Use force=true to re-authenticate.'
            }]
          };
        }
      }

      // Clear tokens if forcing
      if (force) {
        await ts.clearTokens();
      }

      const deviceCode = await ts.initiateDeviceCodeFlow();

      // Start polling in the background — it will save tokens when complete
      const pollPromise = ts.pollForDeviceCodeToken(
        deviceCode.device_code,
        deviceCode.interval || 5,
        deviceCode.expires_in || 900
      );

      // Store the poll promise so check-auth-status can report on it
      ts._pendingDeviceCodePoll = pollPromise;
      pollPromise
        .then(() => { ts._pendingDeviceCodePoll = null; })
        .catch(() => { ts._pendingDeviceCodePoll = null; });

      return {
        content: [{
          type: "text",
          text: `To authenticate, open your browser and go to:\n\n  ${deviceCode.verification_uri}\n\nEnter the code: ${deviceCode.user_code}\n\nThis code expires in ${Math.floor((deviceCode.expires_in || 900) / 60)} minutes.\n\nOnce you've completed the sign-in, use 'check-auth-status' to verify.`
        }]
      };
    } catch (error) {
      return {
        content: [{
          type: "text",
          text: `Device code authentication failed: ${error.message}`
        }]
      };
    }
  }

  // Authorization code flow (default) — generate auth URL
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
  const ts = getTokenStorage();

  try {
    const accessToken = await ts.getValidAccessToken();
    if (accessToken) {
      const expiresAt = ts.getExpiryTime();
      const remainingMin = Math.round((expiresAt - Date.now()) / 60000);
      return {
        content: [{ type: "text", text: `Authenticated and ready. Token valid for ~${remainingMin} minutes.` }]
      };
    }
  } catch {
    // fall through
  }

  // Check if device code poll is in progress
  if (ts._pendingDeviceCodePoll) {
    return {
      content: [{ type: "text", text: "Waiting for you to complete device code sign-in..." }]
    };
  }

  return {
    content: [{ type: "text", text: "Not authenticated. Use the 'authenticate' tool to start." }]
  };
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
