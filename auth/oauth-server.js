const express = require('express');
const querystring = require('querystring');
const https = require('https');
const fs = require('fs');
const crypto = require('crypto');
const TokenStorage = require('./token-storage');
const { escapeHtml } = require('../utils/html-sanitizer');

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
  }

  app.get('/auth', (req, res) => {
    if (!authConfig.clientId) {
      return res.status(500).send(templates.authError('Configuration Error', 'Client ID is not configured.'));
    }
    // Generate CSRF state -- consuming application must store and validate this value
    const state = crypto.randomBytes(16).toString('hex');

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

    // Reject requests missing the state parameter to prevent CSRF attacks.
    // Full state value validation is the responsibility of the consuming application,
    // since this module does not manage sessions.
    if (!state) {
      console.error("OAuth callback received without a 'state' parameter.");
      return res.status(400).send(templates.authError('Missing State Parameter', 'The state parameter was missing from the OAuth callback. Please try authenticating again.'));
    }

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
};
