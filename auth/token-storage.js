const fs = require('fs').promises;
const path = require('path');
const https = require('https');
const querystring = require('querystring');

class TokenStorage {
  constructor(config) {
    this.config = {
      tokenStorePath: path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json'),
      clientId: process.env.MS_CLIENT_ID,
      clientSecret: process.env.MS_CLIENT_SECRET,
      redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
      scopes: (process.env.MS_SCOPES || 'offline_access User.Read Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Contacts.Read').split(' '),
      tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      refreshTokenBuffer: 5 * 60 * 1000, // 5 minutes buffer for token refresh
      ...config // Allow overriding default config
    };
    this.tokens = null;
    this._loadPromise = null;
    this._refreshPromise = null;

    if (!this.config.clientId || !this.config.clientSecret) {
      console.warn("TokenStorage: MS_CLIENT_ID or MS_CLIENT_SECRET is not configured. Token operations might fail.");
    }
  }

  async _loadTokensFromFile() {
    try {
      const tokenData = await fs.readFile(this.config.tokenStorePath, 'utf8');
      this.tokens = JSON.parse(tokenData);
      console.log('Tokens loaded from file.');
      return this.tokens;
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.log('Token file not found. No tokens loaded.');
      } else {
        console.error('Error loading token cache:', error);
      }
      this.tokens = null;
      return null;
    }
  }

  async _saveTokensToFile() {
    if (!this.tokens) {
      console.warn('No tokens to save.');
      return false;
    }
    try {
      await fs.writeFile(this.config.tokenStorePath, JSON.stringify(this.tokens, null, 2));
      console.log('Tokens saved successfully.');
      // return true; // No longer returning boolean, will throw on error.
    } catch (error) {
      console.error('Error saving token cache:', error);
      throw error; // Propagate the error
    }
  }

  async getTokens() {
    if (this.tokens) {
      return this.tokens;
    }
    if (!this._loadPromise) {
        this._loadPromise = this._loadTokensFromFile().finally(() => {
            this._loadPromise = null; // Reset promise once completed
        });
    }
    return this._loadPromise;
  }

  // Get all stored accounts (for multi-account support)
  async getAllAccounts() {
    const allTokens = await this.getTokens();
    if (!allTokens) return [];

    // Handle new format with accounts object
    if (allTokens.accounts && typeof allTokens.accounts === 'object') {
      return Object.keys(allTokens.accounts);
    }

    // Handle old format with single token (backward compatibility)
    if (allTokens.access_token) {
      return ['default'];
    }

    return [];
  }

  // Set active account for multi-account support
  async setActiveAccount(email) {
    const allTokens = await this.getTokens();
    if (!allTokens || !allTokens.accounts || !allTokens.accounts[email]) {
      throw new Error(`Account ${email} not found`);
    }
    if (!allTokens.accounts) allTokens.accounts = {};
    allTokens.activeAccount = email;
    this.tokens = allTokens;
    await this._saveTokensToFile();
    console.log(`Active account switched to: ${email}`);
  }

  // Get active account email
  async getActiveAccount() {
    const allTokens = await this.getTokens();
    if (!allTokens) return null;
    return allTokens.activeAccount || 'default';
  }

  getExpiryTime() {
    return this.tokens && this.tokens.expires_at ? this.tokens.expires_at : 0;
  }

  isTokenExpired() {
    if (!this.tokens || !this.tokens.expires_at) {
      return true; // No token or no expiry means it's effectively expired or invalid
    }
    // Check if current time is past expiry time, considering a buffer
    return Date.now() >= (this.tokens.expires_at - this.config.refreshTokenBuffer);
  }

  async getValidAccessToken() {
    const allTokens = await this.getTokens(); // Ensure tokens are loaded

    if (!allTokens) {
      console.log('No tokens available.');
      return null;
    }

    // Handle new multi-account format
    if (allTokens.accounts && typeof allTokens.accounts === 'object') {
      const activeAccount = allTokens.activeAccount;
      if (!activeAccount || !allTokens.accounts[activeAccount]) {
        console.log(`No valid active account. Available accounts: ${Object.keys(allTokens.accounts).join(', ')}`);
        return null;
      }

      const activeTokens = allTokens.accounts[activeAccount];
      if (!activeTokens.access_token) {
        console.log(`No access token for active account: ${activeAccount}`);
        return null;
      }

      // Check expiration
      const isExpired = Date.now() >= (activeTokens.expires_at - this.config.refreshTokenBuffer);
      if (isExpired) {
        console.log(`Access token for ${activeAccount} expired or nearing expiration. Attempting refresh.`);
        if (activeTokens.refresh_token) {
          try {
            return await this.refreshAccessToken();
          } catch (refreshError) {
            console.error('Failed to refresh access token:', refreshError);
            return null;
          }
        } else {
          console.warn('No refresh token available for this account.');
          return null;
        }
      }

      return activeTokens.access_token;
    }

    // Handle old single-token format (backward compatibility)
    if (!this.tokens || !this.tokens.access_token) {
      console.log('No access token available.');
      return null;
    }

    if (this.isTokenExpired()) {
      console.log('Access token expired or nearing expiration. Attempting refresh.');
      if (this.tokens.refresh_token) {
        try {
          return await this.refreshAccessToken();
        } catch (refreshError) {
          console.error('Failed to refresh access token:', refreshError);
          this.tokens = null; // Invalidate tokens on refresh failure
          await this._saveTokensToFile(); // Persist invalidation
          return null;
        }
      } else {
        console.warn('No refresh token available. Cannot refresh access token.');
        this.tokens = null; // Invalidate tokens as they are expired and cannot be refreshed
        await this._saveTokensToFile(); // Persist invalidation
        return null;
      }
    }
    return this.tokens.access_token;
  }

  async refreshAccessToken() {
    if (!this.tokens || !this.tokens.refresh_token) {
      throw new Error('No refresh token available to refresh the access token.');
    }

    // Prevent multiple concurrent refresh attempts
    if (this._refreshPromise) {
        console.log("Refresh already in progress, returning existing promise.");
        return this._refreshPromise.then(tokens => tokens.access_token);
    }

    console.log('Attempting to refresh access token...');
    const postData = querystring.stringify({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: this.tokens.refresh_token,
      scope: this.config.scopes.join(' ')
    });

    const requestOptions = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    this._refreshPromise = new Promise((resolve, reject) => {
        const req = https.request(this.config.tokenEndpoint, requestOptions, (res) => {
            let data = '';
            res.on('data', (chunk) => data += chunk);
            res.on('end', async () => {
                try {
                    const responseBody = JSON.parse(data);
                    if (res.statusCode >= 200 && res.statusCode < 300) {
                        this.tokens.access_token = responseBody.access_token;
                        // Microsoft Graph API refresh tokens may or may not return a new refresh_token
                        if (responseBody.refresh_token) {
                            this.tokens.refresh_token = responseBody.refresh_token;
                        }
                        this.tokens.expires_in = responseBody.expires_in;
                        this.tokens.expires_at = Date.now() + (responseBody.expires_in * 1000);
                        try {
                            await this._saveTokensToFile();
                            console.log('Access token refreshed and saved successfully.');
                            resolve(this.tokens);
                        } catch (saveError) {
                            console.error('Failed to save refreshed tokens:', saveError);
                            // Even if save fails, tokens are updated in memory.
                            // Depending on desired strictness, could reject here.
                            // For now, resolve with in-memory tokens but log critical error.
                            // Or, to be stricter and align with re-throwing:
                            reject(new Error(`Access token refreshed but failed to save: ${saveError.message}`));
                        }
                    } else {
                        console.error('Error refreshing token:', responseBody);
                        reject(new Error(responseBody.error_description || `Token refresh failed with status ${res.statusCode}`));
                    }
                } catch (e) { // Catch any error during parsing or saving
                    console.error('Error processing refresh token response or saving tokens:', e);
                    reject(e);
                } finally {
                    this._refreshPromise = null; // Clear promise after completion
                }
            });
        });
        req.on('error', (error) => {
            console.error('HTTP error during token refresh:', error);
            reject(error);
            this._refreshPromise = null; // Clear promise on error
        });
        req.write(postData);
        req.end();
    });

    return this._refreshPromise.then(tokens => tokens.access_token);
  }


  async exchangeCodeForTokens(authCode) {
    if (!this.config.clientId || !this.config.clientSecret) {
        throw new Error("Client ID or Client Secret is not configured. Cannot exchange code for tokens.");
    }
    console.log('Exchanging authorization code for tokens...');
    const postData = querystring.stringify({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: this.config.redirectUri,
      scope: this.config.scopes.join(' ')
    });

    const requestOptions = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    return new Promise((resolve, reject) => {
      const req = https.request(this.config.tokenEndpoint, requestOptions, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', async () => {
          try {
            const responseBody = JSON.parse(data);
            if (res.statusCode >= 200 && res.statusCode < 300) {
              const newToken = {
                access_token: responseBody.access_token,
                refresh_token: responseBody.refresh_token,
                expires_in: responseBody.expires_in,
                expires_at: Date.now() + (responseBody.expires_in * 1000),
                scope: responseBody.scope,
                token_type: responseBody.token_type
              };

              // Get user info to determine which account this is for
              try {
                const userEmail = await this._getUserEmail(newToken.access_token);

                // Load existing tokens and add this new account
                let allTokens = await this.getTokens() || {};
                if (!allTokens.accounts) allTokens.accounts = {};

                allTokens.accounts[userEmail] = newToken;
                allTokens.activeAccount = userEmail;

                this.tokens = allTokens;
                await this._saveTokensToFile();
                console.log(`Tokens exchanged and saved successfully for account: ${userEmail}`);
                resolve({...newToken, email: userEmail});
              } catch (emailError) {
                console.warn('Could not determine user email, using timestamp-based key', emailError);
                // Fallback: use timestamp as key if we can't get email
                const accountKey = `account_${Date.now()}`;
                let allTokens = await this.getTokens() || {};
                if (!allTokens.accounts) allTokens.accounts = {};

                allTokens.accounts[accountKey] = newToken;
                allTokens.activeAccount = accountKey;

                this.tokens = allTokens;
                await this._saveTokensToFile();
                resolve({...newToken, email: accountKey});
              }
            } else {
              console.error('Error exchanging code for tokens:', responseBody);
              reject(new Error(responseBody.error_description || `Token exchange failed with status ${res.statusCode}`));
            }
          } catch (e) { // Catch any error during parsing or saving
            console.error('Error processing token exchange response or saving tokens:', e, "Raw data:", data);
            reject(new Error(`Error processing token response: ${e.message}. Response data: ${data}`));
          }
        });
      });
      req.on('error', (error) => {
        console.error('HTTP error during code exchange:', error);
        reject(error);
      });
      req.write(postData);
      req.end();
    });
  }

  // Helper method to get user email from access token
  async _getUserEmail(accessToken) {
    return new Promise((resolve, reject) => {
      const requestOptions = {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      };

      const req = https.request('https://graph.microsoft.com/v1.0/me', requestOptions, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', () => {
          try {
            const userInfo = JSON.parse(data);
            const email = userInfo.mail || userInfo.userPrincipalName;
            if (email) {
              resolve(email);
            } else {
              reject(new Error('No email found in user info'));
            }
          } catch (e) {
            reject(new Error(`Failed to parse user info: ${e.message}`));
          }
        });
      });

      req.on('error', reject);
      req.end();
    });
  }

  // Utility to clear tokens, e.g., for logout or forcing re-auth
  async clearTokens() {
    this.tokens = null;
    try {
      await fs.unlink(this.config.tokenStorePath);
      console.log('Token file deleted successfully.');
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.log('Token file not found, nothing to delete.');
      } else {
        console.error('Error deleting token file:', error);
      }
    }
  }
}

module.exports = TokenStorage;
// Adding a newline at the end of the file as requested by Gemini Code Assist
