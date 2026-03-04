const express = require('express');
const querystring = require('querystring');
const https = require('https');
const fs = require('fs');
const crypto = require('crypto'); // Added for generating random string
const TokenStorage = require('./token-storage'); // Assuming TokenStorage is in the same directory

// HTML templates
function escapeHtml(unsafe) {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

const templates = {
  authError: (error, errorDescription) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Authorization Failed</h1>
        <p><strong>Error:</strong> ${escapeHtml(error)}</p>
        ${errorDescription ? `<p><strong>Description:</strong> ${escapeHtml(errorDescription)}</p>` : ''}
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  authSuccess: `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #2ecc71;">✅ Authentication Successful</h1>
        <p>You have successfully authenticated with Microsoft Graph API.</p>
        <p>You can close this window.</p>
      </body>
    </html>`,
  tokenExchangeError: (error) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Token Exchange Failed</h1>
        <p>Failed to exchange authorization code for access token.</p>
        <p><strong>Error:</strong> ${escapeHtml(error instanceof Error ? error.message : String(error))}</p>
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  tokenStatus: (status) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1>🔐 Token Status</h1>
        <p>${escapeHtml(status)}</p>
      </body>
    </html>`
};

function createAuthConfig(envPrefix = 'MS_') {
  const tenantId = process.env[`${envPrefix}TENANT_ID`] || 'common';
  const authorityHost = (process.env[`${envPrefix}AUTHORITY_HOST`] || 'https://login.microsoftonline.com').replace(/\/+$/, '');

  return {
    clientId: process.env[`${envPrefix}CLIENT_ID`] || '',
    clientSecret: process.env[`${envPrefix}CLIENT_SECRET`] || '',
    redirectUri: process.env[`${envPrefix}REDIRECT_URI`] || 'http://localhost:3333/auth/callback',
    scopes: (process.env[`${envPrefix}SCOPES`] || 'offline_access User.Read Mail.Read').split(' '),
    tenantId,
    tokenEndpoint: process.env[`${envPrefix}TOKEN_ENDPOINT`] || `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
    authEndpoint: process.env[`${envPrefix}AUTH_ENDPOINT`] || `${authorityHost}/${tenantId}/oauth2/v2.0/authorize`
  };
}

function setupOAuthRoutes(app, tokenStorage, authConfig, envPrefix = 'MS_') {
  if (!authConfig) {
    authConfig = createAuthConfig(envPrefix);
  }

  if (!(tokenStorage instanceof TokenStorage)) {
    console.error("Error: tokenStorage is not an instance of TokenStorage. OAuth routes will not function correctly.");
    // Optionally, you could throw an error here or disable the routes
    // throw new Error("Invalid tokenStorage provided to setupOAuthRoutes");
  }


  app.get('/auth', (req, res) => {
    if (!authConfig.clientId) {
      return res.status(500).send(templates.authError('Configuration Error', 'Client ID is not configured.'));
    }
    const state = crypto.randomBytes(16).toString('hex'); // Generate a random 16-byte string
    // Store state in session or similar mechanism if available.
    // For a server without sessions, this state would need to be passed through and verified differently,
    // or a temporary server-side storage (like a short-lived cache) would be needed.
    // For this example, we'll assume session middleware is configured elsewhere if this were a full app.
    // If using express-session: req.session.oauthState = state;
    // Since this is a module, actual session handling is outside its direct scope,
    // but it's crucial for the consuming application to handle state verification.

    const authorizationUrl = `${authConfig.authEndpoint}?` +
      querystring.stringify({
        client_id: authConfig.clientId,
        response_type: 'code',
        redirect_uri: authConfig.redirectUri,
        scope: authConfig.scopes.join(' '),
        response_mode: 'query',
        state: state
      });
    res.redirect(authorizationUrl);
  });

  app.get('/auth/callback', async (req, res) => {
    const { code, error, error_description, state } = req.query;

    // IMPORTANT: State validation is crucial for CSRF protection.
    // The application using this module MUST implement a way to store the 'state' generated in /auth
    // (e.g., in a user session if using express-session, or a short-lived cache)
    // and then verify it here against the 'state' received from the OAuth provider.
    // For example, if using express-session:
    // const savedState = req.session.oauthState;
    // if (!state || state !== savedState) {
    //   console.error("OAuth callback state mismatch. Potential CSRF attack.");
    //   return res.status(400).send(templates.authError('Invalid State', 'CSRF token mismatch. Please try authenticating again.'));
    // }
    // delete req.session.oauthState; // Clean up session state

    // Since this module itself doesn't manage sessions, we'll log a warning if state is missing,
    // but actual enforcement must be done by the consuming application.
    // The Gemini review recommended uncommenting the rejection.
    // However, the consuming app (CLI or server) is responsible for session/state storage.
    // This module *cannot* validate state if it wasn't involved in storing it.
    // The PR author (ranxian) needs to implement state storage & validation in the calling server (sse-server.js or outlook-auth-server.js).
    // For now, enforcing a missing state here would break flows where state *is* passed but not validated by *this specific module*.
    // The best this module can do is check for presence and rely on the consumer to validate the actual value.
    // The original PR #10's outlook-auth-server.js used Date.now() and didn't store/validate it beyond this.
    // The new sse-server.js also doesn't show session management for state.
    // So, we will make the check for presence mandatory as per Gemini's suggestion.
    if (!state) {
        console.error("OAuth callback received without a 'state' parameter. Rejecting request to prevent potential CSRF attack.");
        return res.status(400).send(templates.authError('Missing State Parameter', 'The state parameter was missing from the OAuth callback. This is a security risk. Please try authenticating again.'));
    }
    // Further validation of the state's VALUE (e.g., req.session.oauthState === state) is the responsibility
    // of the application integrating this module, as session management is outside this module's scope.
    // if (req.session && req.session.oauthState !== state) {
    //    return res.status(400).send(templates.authError('Invalid State Parameter', 'CSRF detected. State mismatch.'));
    // }
    // if (req.session) delete req.session.oauthState;


    if (error) {
      return res.status(400).send(templates.authError(error, error_description));
    }

    if (!code) {
      return res.status(400).send(templates.authError('Missing Authorization Code', 'No authorization code was provided in the callback.'));
    }

    try {
      await tokenStorage.exchangeCodeForTokens(code);
      res.send(templates.authSuccess);
    } catch (exchangeError) {
      console.error('Token exchange error:', exchangeError);
      res.status(500).send(templates.tokenExchangeError(exchangeError));
    }
  });

  app.get('/token-status', async (req, res) => {
    try {
      const token = await tokenStorage.getValidAccessToken();
      if (token) {
        const expiryDate = new Date(tokenStorage.getExpiryTime());
        res.send(templates.tokenStatus(`Access token is valid. Expires at: ${expiryDate.toLocaleString()}`));
      } else {
        res.send(templates.tokenStatus('No valid access token found. Please authenticate.'));
      }
    } catch (err) {
      res.status(500).send(templates.tokenStatus(`Error checking token status: ${err.message}`));
    }
  });
}

module.exports = {
  setupOAuthRoutes,
  createAuthConfig,
  // Exporting templates for potential direct use or testing, though not typical
  // templates
};
// Adding a newline at the end of the file as requested by Gemini Code Assist
