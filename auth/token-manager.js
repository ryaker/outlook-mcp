/**
 * Token management for Microsoft Graph API authentication
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
function getFlowAccessToken() {
  const tokens = loadTokenCache();
  if (!tokens) return null;

  // Check if flow token exists and is not expired
  if (tokens.flow_access_token && tokens.flow_expires_at) {
    if (Date.now() < tokens.flow_expires_at) {
      return tokens.flow_access_token;
    }
  }

  return null;
}

/**
 * Saves Flow API tokens alongside existing Graph tokens
 * @param {object} flowTokens - The Flow tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveFlowTokens(flowTokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;

    // Load existing tokens
    let existingTokens = {};
    if (fs.existsSync(tokenPath)) {
      const tokenData = fs.readFileSync(tokenPath, 'utf8');
      existingTokens = JSON.parse(tokenData);
    }

    // Merge flow tokens
    const mergedTokens = {
      ...existingTokens,
      flow_access_token: flowTokens.access_token,
      flow_refresh_token: flowTokens.refresh_token,
      flow_expires_at: flowTokens.expires_at || (Date.now() + (flowTokens.expires_in || 3600) * 1000)
    };

    fs.writeFileSync(tokenPath, JSON.stringify(mergedTokens, null, 2), { mode: 0o600 });
    console.log('Flow tokens saved successfully');

    // Update cache
    cachedTokens = mergedTokens;
    return true;
  } catch (error) {
    console.error('Error saving Flow tokens:', error);
    return false;
  }
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,
  getFlowAccessToken,
  saveFlowTokens,
  createTestTokens
};
