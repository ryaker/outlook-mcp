/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;
let refreshPromise = null;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`[DEBUG] Attempting to load tokens from: ${tokenPath}`);
    console.error(`[DEBUG] HOME directory: ${process.env.HOME}`);
    console.error(`[DEBUG] Full resolved path: ${tokenPath}`);
    
    // Log file existence and details
    if (!fs.existsSync(tokenPath)) {
      console.error('[DEBUG] Token file does not exist');
      return null;
    }
    
    const stats = fs.statSync(tokenPath);
    console.error(`[DEBUG] Token file stats:
      Size: ${stats.size} bytes
      Created: ${stats.birthtime}
      Modified: ${stats.mtime}`);
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    console.error('[DEBUG] Token file contents length:', tokenData.length);
    console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
    
    try {
      const tokens = JSON.parse(tokenData);
      console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
      
      // Log each key's value to see what's present
      Object.keys(tokens).forEach(key => {
        console.error(`[DEBUG] ${key}: ${typeof tokens[key]}`);
      });
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[DEBUG] No access_token found in tokens');
        return null;
      }
      
      // Update the cache
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[DEBUG] Error parsing token JSON:', parseError);
      return null;
    }
  } catch (error) {
    console.error('[DEBUG] Error loading token cache:', error);
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
    console.error(`Saving tokens to: ${tokenPath}`);
    
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.error('Tokens saved successfully');
    
    // Update the cache
    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Refreshes the access token using the refresh token
 * @returns {Promise<string|null>} - The new access token or null if refresh fails
 */
async function refreshAccessToken() {
  // Prevent multiple concurrent refresh attempts
  if (refreshPromise) {
    console.error('[DEBUG] Refresh already in progress, waiting...');
    return refreshPromise;
  }

  const tokens = cachedTokens || loadTokenCache();
  
  if (!tokens || !tokens.refresh_token) {
    console.error('[DEBUG] No refresh token available');
    return null;
  }

  console.error('[DEBUG] Attempting to refresh access token...');
  
  refreshPromise = new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: config.AUTH_CONFIG.clientId,
      client_secret: config.AUTH_CONFIG.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: tokens.refresh_token,
      scope: config.AUTH_CONFIG.scopes.join(' ')
    });

    const tenantId = process.env.MS_TENANT_ID || 'common';
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          const responseBody = JSON.parse(data);
          if (res.statusCode >= 200 && res.statusCode < 300) {
            // Update tokens
            tokens.access_token = responseBody.access_token;
            // Microsoft may or may not return a new refresh token
            if (responseBody.refresh_token) {
              tokens.refresh_token = responseBody.refresh_token;
            }
            tokens.expires_in = responseBody.expires_in;
            tokens.expires_at = Date.now() + (responseBody.expires_in * 1000);
            
            // Save updated tokens
            if (saveTokenCache(tokens)) {
              console.error('[DEBUG] Access token refreshed and saved successfully');
              resolve(tokens.access_token);
            } else {
              console.error('[DEBUG] Failed to save refreshed tokens');
              resolve(null);
            }
          } else {
            console.error('[DEBUG] Token refresh failed:', responseBody);
            resolve(null);
          }
        } catch (e) {
          console.error('[DEBUG] Error processing refresh response:', e);
          resolve(null);
        } finally {
          refreshPromise = null;
        }
      });
    });

    req.on('error', (error) => {
      console.error('[DEBUG] HTTP error during token refresh:', error);
      refreshPromise = null;
      resolve(null);
    });

    req.write(postData);
    req.end();
  });

  return refreshPromise;
}

/**
 * Gets the current access token, refreshing if necessary
 * @returns {Promise<string|null>} - The access token or null if not available
 */
async function getAccessToken() {
  // Load tokens if not cached
  if (!cachedTokens) {
    const tokens = loadTokenCache();
    if (!tokens) {
      return null;
    }
  }

  // Check token expiration with 5 minute buffer
  const now = Date.now();
  const expiresAt = cachedTokens.expires_at || 0;
  const refreshBuffer = 5 * 60 * 1000; // 5 minutes

  console.error(`[DEBUG] Current time: ${now}`);
  console.error(`[DEBUG] Token expires at: ${expiresAt}`);

  // If token is expired or will expire soon, try to refresh
  if (now >= (expiresAt - refreshBuffer)) {
    console.error('[DEBUG] Token expired or expiring soon, attempting refresh...');
    const newToken = await refreshAccessToken();
    return newToken;
  }

  return cachedTokens.access_token;
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
  createTestTokens,
  refreshAccessToken
};
