/**
 * Enhanced token storage with automatic refresh capabilities
 * Implements Microsoft OAuth 2.0 refresh token flow
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

class TokenStorage {
  constructor() {
    this.tokenPath = config.AUTH_CONFIG.tokenStorePath;
    this.clientId = process.env.MS_CLIENT_ID || '';
    this.clientSecret = process.env.MS_CLIENT_SECRET || '';
    this.scopes = [
      'offline_access',
      'User.Read', 
      'Mail.Read',
      'Mail.Send',
      'Calendars.Read',
      'Calendars.ReadWrite',
      'Contacts.Read'
    ];
    this.cachedTokens = null;
    
    // Buffer time before expiration to refresh token (5 minutes)
    this.refreshBufferTime = 5 * 60 * 1000;
  }

  /**
   * Gets a valid access token, refreshing if necessary
   * @returns {Promise<string|null>} Valid access token or null if authentication required
   */
  async getValidAccessToken() {
    try {
      const tokens = await this.loadTokens();
      if (!tokens) {
        console.error('[TokenStorage] No tokens available');
        return null;
      }

      // Check if token needs refresh
      if (this.needsRefresh(tokens)) {
        console.error('[TokenStorage] Token needs refresh');
        const refreshedTokens = await this.refreshToken(tokens.refresh_token);
        if (refreshedTokens) {
          return refreshedTokens.access_token;
        } else {
          console.error('[TokenStorage] Token refresh failed');
          return null;
        }
      }

      return tokens.access_token;
    } catch (error) {
      console.error('[TokenStorage] Error getting valid access token:', error.message);
      return null;
    }
  }

  /**
   * Loads tokens from storage
   * @returns {Promise<object|null>} Token object or null
   */
  async loadTokens() {
    try {
      if (this.cachedTokens) {
        return this.cachedTokens;
      }

      if (!fs.existsSync(this.tokenPath)) {
        return null;
      }

      const tokenData = fs.readFileSync(this.tokenPath, 'utf8');
      const tokens = JSON.parse(tokenData);

      if (!tokens.access_token) {
        return null;
      }

      this.cachedTokens = tokens;
      return tokens;
    } catch (error) {
      console.error('[TokenStorage] Error loading tokens:', error.message);
      return null;
    }
  }

  /**
   * Saves tokens to storage
   * @param {object} tokens - Token object to save
   * @returns {Promise<boolean>} Success status
   */
  async saveTokens(tokens) {
    try {
      // Calculate expiration time if not present
      if (!tokens.expires_at && tokens.expires_in) {
        tokens.expires_at = Date.now() + (tokens.expires_in * 1000);
      }

      fs.writeFileSync(this.tokenPath, JSON.stringify(tokens, null, 2), 'utf8');
      this.cachedTokens = tokens;
      console.error('[TokenStorage] Tokens saved successfully');
      return true;
    } catch (error) {
      console.error('[TokenStorage] Error saving tokens:', error.message);
      return false;
    }
  }

  /**
   * Checks if token needs refresh
   * @param {object} tokens - Token object
   * @returns {boolean} True if refresh is needed
   */
  needsRefresh(tokens) {
    if (!tokens.expires_at) {
      return true; // No expiration info, assume needs refresh
    }

    const now = Date.now();
    const timeUntilExpiry = tokens.expires_at - now;
    
    // Refresh if expires within buffer time
    return timeUntilExpiry <= this.refreshBufferTime;
  }

  /**
   * Refreshes access token using refresh token
   * @param {string} refreshToken - The refresh token
   * @returns {Promise<object|null>} New token object or null if failed
   */
  async refreshToken(refreshToken) {
    if (!refreshToken) {
      console.error('[TokenStorage] No refresh token available');
      return null;
    }

    if (!this.clientId || !this.clientSecret) {
      console.error('[TokenStorage] Client credentials not configured');
      return null;
    }

    return new Promise((resolve, reject) => {
      const postData = querystring.stringify({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: this.scopes.join(' ')
      });

      const options = {
        hostname: 'login.microsoftonline.com',
        path: '/common/oauth2/v2.0/token',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      console.error('[TokenStorage] Refreshing access token...');

      const req = https.request(options, (res) => {
        let data = '';

        res.on('data', (chunk) => {
          data += chunk;
        });

        res.on('end', async () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            try {
              const tokenResponse = JSON.parse(data);
              
              // Calculate expiration time
              const expiresAt = Date.now() + (tokenResponse.expires_in * 1000);
              tokenResponse.expires_at = expiresAt;

              // Preserve refresh token if not provided in response
              if (!tokenResponse.refresh_token) {
                tokenResponse.refresh_token = refreshToken;
              }

              // Save the new tokens
              await this.saveTokens(tokenResponse);
              
              console.error('[TokenStorage] Token refreshed successfully');
              resolve(tokenResponse);
            } catch (error) {
              console.error('[TokenStorage] Error parsing refresh response:', error.message);
              resolve(null);
            }
          } else {
            console.error(`[TokenStorage] Token refresh failed with status ${res.statusCode}: ${data}`);
            resolve(null);
          }
        });
      });

      req.on('error', (error) => {
        console.error('[TokenStorage] Network error during token refresh:', error.message);
        resolve(null);
      });

      req.write(postData);
      req.end();
    });
  }

  /**
   * Exchanges authorization code for tokens
   * @param {string} code - Authorization code
   * @param {string} redirectUri - Redirect URI used in auth request
   * @returns {Promise<object|null>} Token object or null if failed
   */
  async exchangeCodeForTokens(code, redirectUri) {
    if (!this.clientId || !this.clientSecret) {
      throw new Error('Client credentials not configured');
    }

    return new Promise((resolve, reject) => {
      const postData = querystring.stringify({
        client_id: this.clientId,
        client_secret: this.clientSecret,
        code: code,
        redirect_uri: redirectUri,
        grant_type: 'authorization_code',
        scope: this.scopes.join(' ')
      });

      const options = {
        hostname: 'login.microsoftonline.com',
        path: '/common/oauth2/v2.0/token',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      console.error('[TokenStorage] Exchanging code for tokens...');

      const req = https.request(options, (res) => {
        let data = '';

        res.on('data', (chunk) => {
          data += chunk;
        });

        res.on('end', async () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            try {
              const tokenResponse = JSON.parse(data);
              
              // Calculate expiration time
              const expiresAt = Date.now() + (tokenResponse.expires_in * 1000);
              tokenResponse.expires_at = expiresAt;

              // Save tokens
              await this.saveTokens(tokenResponse);
              
              console.error('[TokenStorage] Tokens obtained and saved successfully');
              resolve(tokenResponse);
            } catch (error) {
              reject(new Error(`Error parsing token response: ${error.message}`));
            }
          } else {
            reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
          }
        });
      });

      req.on('error', (error) => {
        reject(new Error(`Network error during token exchange: ${error.message}`));
      });

      req.write(postData);
      req.end();
    });
  }

  /**
   * Clears stored tokens
   */
  clearTokens() {
    try {
      if (fs.existsSync(this.tokenPath)) {
        fs.unlinkSync(this.tokenPath);
      }
      this.cachedTokens = null;
      console.error('[TokenStorage] Tokens cleared');
    } catch (error) {
      console.error('[TokenStorage] Error clearing tokens:', error.message);
    }
  }

  /**
   * Gets token expiration info
   * @returns {Promise<object|null>} Expiration info or null
   */
  async getTokenExpirationInfo() {
    const tokens = await this.loadTokens();
    if (!tokens || !tokens.expires_at) {
      return null;
    }

    const now = Date.now();
    const expiresAt = tokens.expires_at;
    const timeUntilExpiry = expiresAt - now;

    return {
      expiresAt: new Date(expiresAt),
      timeUntilExpiry: timeUntilExpiry,
      isExpired: timeUntilExpiry <= 0,
      needsRefresh: this.needsRefresh(tokens)
    };
  }
}

module.exports = TokenStorage;