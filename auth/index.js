/**
 * Authentication module for Outlook MCP server
 */
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Create global token storage instance
const tokenStorage = new TokenStorage();

/**
 * Ensures the user is authenticated and returns an access token
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    // Force re-authentication
    throw new Error('Authentication required');
  }

  // Try to get a valid access token (with auto-refresh)
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

/**
 * Gets the token storage instance
 * @returns {TokenStorage} - Token storage instance
 */
function getTokenStorage() {
  return tokenStorage;
}

module.exports = {
  tokenStorage,
  authTools,
  ensureAuthenticated,
  getTokenStorage
};
