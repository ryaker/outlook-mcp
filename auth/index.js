/**
 * Authentication module for Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Singleton TokenStorage instance — uses AUTH_CONFIG scopes for device code flow
// and token refresh to request the correct permissions.
const tokenStorage = new TokenStorage({
  clientId: config.AUTH_CONFIG.clientId,
  clientSecret: config.AUTH_CONFIG.clientSecret,
  scopes: config.AUTH_CONFIG.scopes,
  tokenStorePath: config.AUTH_CONFIG.tokenStorePath,
});

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  // Use TokenStorage which handles automatic refresh
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

module.exports = {
  tokenManager,
  tokenStorage,
  authTools,
  ensureAuthenticated
};
