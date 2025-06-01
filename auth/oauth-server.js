/**
 * Shared OAuth server endpoints for Microsoft Graph API authentication
 * Provides reusable Express routes for OAuth flow
 */

/**
 * OAuth HTML templates
 */
const HTML_TEMPLATES = {
  authError: (error, errorDescription) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Authorization Failed</h1>
        <p><strong>Error:</strong> ${error}</p>
        ${errorDescription ? `<p><strong>Description:</strong> ${errorDescription}</p>` : ''}
        <p>You can close this window and try again.</p>
      </body>
    </html>
  `,

  missingCode: () => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Missing Authorization Code</h1>
        <p>No authorization code received from the OAuth provider.</p>
        <p>You can close this window and try again.</p>
      </body>
    </html>
  `,

  authSuccess: () => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #27ae60;">✅ Authorization Successful</h1>
        <p>Your access token has been saved successfully!</p>
        <p>You can close this window and continue using the MCP server.</p>
        <script>
          setTimeout(() => window.close(), 3000);
        </script>
      </body>
    </html>
  `,

  tokenExchangeError: (error) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ Token Exchange Failed</h1>
        <p>Failed to exchange authorization code for access token.</p>
        <p><strong>Error:</strong> ${error instanceof Error ? error.message : 'Unknown error'}</p>
        <p>You can close this window and try again.</p>
      </body>
    </html>
  `,

  oauthNotConfigured: () => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">❌ OAuth Not Configured</h1>
        <p>Microsoft Graph API credentials are not properly configured. Please check your environment variables.</p>
        <p>You can close this window.</p>
      </body>
    </html>
  `
};

/**
 * Setup OAuth routes on an Express app
 * @param {Express} app - Express application instance
 * @param {Object} authConfig - OAuth configuration
 * @param {Object} tokenStorage - Token storage instance
 */
function setupOAuthRoutes(app, authConfig, tokenStorage) {
  // OAuth callback endpoint
  app.get('/auth/callback', async (req, res) => {
    const code = req.query.code;
    const error = req.query.error;
    const errorDescription = req.query.error_description;

    if (error) {
      console.error(`Authentication error: ${error} - ${errorDescription}`);
      res.status(400).send(HTML_TEMPLATES.authError(error, errorDescription));
      return;
    }

    if (!code) {
      console.error('No authorization code provided');
      res.status(400).send(HTML_TEMPLATES.missingCode());
      return;
    }

    try {
      console.error('Authorization code received, exchanging for tokens...');
      const tokens = await tokenStorage.exchangeCodeForTokens(code, authConfig.redirectUri);
      console.error('Token exchange successful');
      res.status(200).send(HTML_TEMPLATES.authSuccess());
    } catch (error) {
      console.error('Token exchange failed:', error);
      res.status(500).send(HTML_TEMPLATES.tokenExchangeError(error));
    }
  });

  // OAuth initiation endpoint
  app.get('/auth', async (req, res) => {
    console.error('Auth request received, redirecting to Microsoft login...');

    if (!authConfig.clientId || !authConfig.clientSecret) {
      res.status(500).send(HTML_TEMPLATES.oauthNotConfigured());
      return;
    }

    const clientId = req.query.client_id || authConfig.clientId;

    const authParams = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: authConfig.redirectUri,
      scope: authConfig.scopes.join(' '),
      response_mode: 'query',
      state: Date.now().toString()
    });

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${authParams.toString()}`;
    console.error(`Redirecting to: ${authUrl}`);

    res.redirect(authUrl);
  });

  // Token status endpoint
  app.get('/token-status', async (req, res) => {
    try {
      const tokenInfo = await tokenStorage.getTokenExpirationInfo();
      res.json({
        hasToken: !!tokenInfo,
        tokenInfo: tokenInfo
      });
    } catch (error) {
      res.status(500).json({
        error: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  });
}

/**
 * Create standard auth configuration from environment variables
 * @param {number} port - Server port for redirect URI
 * @returns {Object} Auth configuration
 */
function createAuthConfig(port = 3333) {
  return {
    clientId: process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.MS_CLIENT_SECRET || '',
    redirectUri: `http://localhost:${port}/auth/callback`,
    scopes: [
      'offline_access',
      'User.Read',
      'Mail.Read',
      'Mail.Send',
      'Calendars.Read',
      'Calendars.ReadWrite',
      'Contacts.Read'
    ]
  };
}

module.exports = {
  setupOAuthRoutes,
  createAuthConfig,
  HTML_TEMPLATES
};