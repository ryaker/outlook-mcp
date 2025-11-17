/**
 * Authentication module for Outlook MCP server
 */
const config = require('../config');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Initialize TokenStorage for multi-account support
const tokenStorage = new TokenStorage(config.AUTH_CONFIG);

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

  // Check for existing token using new TokenStorage (multi-account support)
  try {
    const accessToken = await tokenStorage.getValidAccessToken();
    if (!accessToken) {
      throw new Error('Authentication required');
    }
    return accessToken;
  } catch (error) {
    throw new Error('Authentication required');
  }
}

module.exports = {
  authTools,
  ensureAuthenticated,
  tokenStorage
};
