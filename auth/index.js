/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const { authTools } = require('./tools');

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
  
  // Check for existing token and refresh if needed
  const accessToken = await tokenManager.getAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }
  
  return accessToken;
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
