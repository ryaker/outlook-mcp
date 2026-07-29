/**
 * @deprecated Use TokenStorage (auth/token-storage.js) instead.
 * TokenStorage provides automatic refresh, unified scope management,
 * and async file operations. This module is retained only for:
 *   - Test mode token creation (createTestTokens)
 *   - Backwards compatibility for legacy imports
 * Do not add new functionality here. Migrate consumers to TokenStorage.
 */
const fs = require('fs');
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;

    if (!fs.existsSync(tokenPath)) {
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, 'utf8');

    try {
      const tokens = JSON.parse(tokenData);

      if (!tokens.access_token) {
        return null;
      }

      // Check token expiration
      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;

      if (now > expiresAt) {
        return null;
      }

      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('Error parsing token file:', parseError.message);
      return null;
    }
  } catch (error) {
    console.error('Error loading token cache:', error.message);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;

    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2), { mode: 0o600 });

    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error.message);
    return false;
  }
}

/**
 * Gets the current Graph API access token, loading from cache if necessary
 * @returns {string|null} - The access token or null if not available
 */
function getAccessToken() {
  if (cachedTokens && cachedTokens.access_token) {
    return cachedTokens.access_token;
  }

  const tokens = loadTokenCache();
  return tokens ? tokens.access_token : null;
}

/**
 * Gets the current Flow API access token
 * @returns {string|null} - The Flow access token or null if not available
 */
/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: 'test_access_token_' + Date.now(),
    refresh_token: 'test_refresh_token_' + Date.now(),
    expires_at: Date.now() + 3600 * 1000, // 1 hour
  };

  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  createTestTokens,
};
