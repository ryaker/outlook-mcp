const express = require('express');
const request = require('supertest');
const { setupOAuthRoutes, createAuthConfig } = require('../../auth/oauth-server');
const TokenStorage = require('../../auth/token-storage');

jest.mock('../../auth/token-storage'); // Mock TokenStorage class

const mockAuthConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  redirectUri: 'http://localhost:3334/auth/callback', // Use a different port or path for testing
  scopes: ['test_scope', 'openid'],
  tokenEndpoint: 'https://login.example.com/token',
  authEndpoint: 'https://login.example.com/authorize',
};

// Provide a minimal mock for TokenStorage instance methods
const mockTokenStorageInstance = {
  exchangeCodeForTokens: jest.fn(),
  getValidAccessToken: jest.fn(),
  getExpiryTime: jest.fn(),
};

describe('OAuth Server Routes', () => {
  let app;

  beforeEach(() => {
    jest.clearAllMocks();
    TokenStorage.mockImplementation(() => mockTokenStorageInstance); // Return the mocked instance

    app = express();
    // It's important that the TokenStorage instance passed to setupOAuthRoutes is the one we can control.
    // For these tests, we don't need a real TokenStorage, so we pass the mock.
    setupOAuthRoutes(app, mockTokenStorageInstance, mockAuthConfig);
  });

  describe('GET /auth', () => {
    it('should redirect to the OAuth provider with correct parameters', async () => {
      const response = await request(app).get('/auth');
      expect(response.status).toBe(302); // Redirect

      const redirectUrl = new URL(response.headers.location);
      expect(redirectUrl.origin).toBe(mockAuthConfig.authEndpoint.split('/authorize')[0]);
      expect(redirectUrl.pathname).toBe('/authorize');
      expect(redirectUrl.searchParams.get('client_id')).toBe(mockAuthConfig.clientId);
      expect(redirectUrl.searchParams.get('response_type')).toBe('code');
      expect(redirectUrl.searchParams.get('redirect_uri')).toBe(mockAuthConfig.redirectUri);
      expect(redirectUrl.searchParams.get('scope')).toBe(mockAuthConfig.scopes.join(' '));
      expect(redirectUrl.searchParams.get('response_mode')).toBe('query');
      expect(redirectUrl.searchParams.get('state')).toBeDefined();
      expect(redirectUrl.searchParams.get('state').length).toBe(32); // crypto.randomBytes(16).toString('hex')
    });

    it('should return 500 if clientId is not configured', async () => {
      const tempApp = express();
      // Create a new authConfig without clientId for this specific test
      const noClientIdAuthConfig = { ...mockAuthConfig, clientId: null };
      setupOAuthRoutes(tempApp, mockTokenStorageInstance, noClientIdAuthConfig);

      const response = await request(tempApp).get('/auth');
      expect(response.status).toBe(500);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Error:</strong> Configuration Error');
      expect(response.text).toContain('Description:</strong> Client ID is not configured.');
    });
  });

  describe('GET /auth/callback', () => {
    const mockAuthCode = 'mock_auth_code';
    const mockState = 'mock_state_value'; // Example state

    it('should exchange code for tokens and return success HTML', async () => {
      mockTokenStorageInstance.exchangeCodeForTokens.mockResolvedValue({ access_token: 'mock_access_token' });

      // Note: State validation is mocked/skipped here as session management is outside this module.
      // In a real app, the 'state' would be generated in /auth, stored (e.g. session),
      // and verified here. The test passes 'state' to simulate it coming from provider.
      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}&state=${mockState}`);

      expect(mockTokenStorageInstance.exchangeCodeForTokens).toHaveBeenCalledWith(mockAuthCode);
      expect(response.status).toBe(200);
      expect(response.text).toContain('Authentication Successful');
    });

    it('should return 400 and error HTML if OAuth provider returns an error', async () => {
      const oauthError = 'access_denied';
      const oauthErrorDesc = 'User denied access';
      const response = await request(app).get(`/auth/callback?error=${oauthError}&error_description=${oauthErrorDesc}&state=${mockState}`);

      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain(`Error:</strong> ${oauthError}`);
      expect(response.text).toContain(`Description:</strong> ${oauthErrorDesc}`);
      expect(mockTokenStorageInstance.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('should return 400 if no code is provided', async () => {
      const response = await request(app).get(`/auth/callback?state=${mockState}`);
      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Error:</strong> Missing Authorization Code');
      expect(mockTokenStorageInstance.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    // Test for state missing - module currently warns but doesn't block.
    // This test verifies the warning behavior.
    // it('should log a warning if state is missing (but still proceed as per current module logic)', async () => {
    //     const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();
    //     mockTokenStorageInstance.exchangeCodeForTokens.mockResolvedValue({ access_token: 'mock_access_token' });

    //     await request(app).get(`/auth/callback?code=${mockAuthCode}`); // No state passed

    //     expect(consoleWarnSpy).toHaveBeenCalledWith("OAuth callback received without a 'state' parameter. CSRF validation cannot be performed by this module alone.");
    //     expect(mockTokenStorageInstance.exchangeCodeForTokens).toHaveBeenCalledWith(mockAuthCode); // Still called
    //     consoleWarnSpy.mockRestore();
    // });

    it('should return 400 if state is missing from callback', async () => {
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}`); // No state

      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Error:</strong> Missing State Parameter');
      expect(response.text).toContain('The state parameter was missing from the OAuth callback.');
      expect(mockTokenStorageInstance.exchangeCodeForTokens).not.toHaveBeenCalled();
      expect(consoleErrorSpy).toHaveBeenCalledWith("OAuth callback received without a 'state' parameter. Rejecting request to prevent potential CSRF attack.");
      consoleErrorSpy.mockRestore();
    });


    it('should return 500 if token exchange fails', async () => {
      const exchangeError = new Error('Token exchange process failed');
      mockTokenStorageInstance.exchangeCodeForTokens.mockRejectedValue(exchangeError);

      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}&state=${mockState}`);

      expect(response.status).toBe(500);
      expect(response.text).toContain('Token Exchange Failed');
      expect(response.text).toContain(`Error:</strong> ${exchangeError.message}`);
    });
  });

  describe('GET /token-status', () => {
    it('should return valid status if token exists and is valid', async () => {
      const mockAccessToken = 'valid_token_123';
      const mockExpiry = Date.now() + 3600000; // 1 hour from now
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(mockAccessToken);
      mockTokenStorageInstance.getExpiryTime.mockReturnValue(mockExpiry);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(200);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain('Access token is valid.');
      expect(response.text).toContain(`Expires at: ${new Date(mockExpiry).toLocaleString()}`);
    });

    it('should return "no valid token" status if token is not found', async () => {
      mockTokenStorageInstance.getValidAccessToken.mockResolvedValue(null);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(200);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain('No valid access token found. Please authenticate.');
    });

    it('should return 500 if checking token status throws an error', async () => {
      const statusError = new Error('Failed to check token status');
      mockTokenStorageInstance.getValidAccessToken.mockRejectedValue(statusError);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(500);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain(`Error checking token status: ${statusError.message}`);
    });
  });

  describe('createAuthConfig', () => {
    const originalEnv = process.env;

    beforeEach(() => {
        jest.resetModules(); // Important to clear module cache for process.env changes
        process.env = { ...originalEnv }; // Make a copy
    });

    afterAll(() => {
        process.env = originalEnv; // Restore original environment
    });

    it('should use default values if environment variables are not set', () => {
        const config = createAuthConfig('TEST_PREFIX_'); // Use a prefix to avoid collision
        expect(config.clientId).toBe('');
        expect(config.clientSecret).toBe('');
        expect(config.redirectUri).toBe('http://localhost:3333/auth/callback');
        expect(config.scopes).toEqual(['offline_access', 'User.Read', 'Mail.Read']);
        expect(config.tokenEndpoint).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/token');
        expect(config.authEndpoint).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
    });

    it('should use environment variables with specified prefix', () => {
        process.env.MYAPP_CLIENT_ID = 'env_client_id';
        process.env.MYAPP_CLIENT_SECRET = 'env_client_secret';
        process.env.MYAPP_REDIRECT_URI = 'http://env.redirect/uri';
        process.env.MYAPP_SCOPES = 'scope1 scope2';
        process.env.MYAPP_TOKEN_ENDPOINT = 'http://env.token/endpoint';
        process.env.MYAPP_AUTH_ENDPOINT = 'http://env.auth/endpoint';

        const config = createAuthConfig('MYAPP_');

        expect(config.clientId).toBe('env_client_id');
        expect(config.clientSecret).toBe('env_client_secret');
        expect(config.redirectUri).toBe('http://env.redirect/uri');
        expect(config.scopes).toEqual(['scope1', 'scope2']);
        expect(config.tokenEndpoint).toBe('http://env.token/endpoint');
        expect(config.authEndpoint).toBe('http://env.auth/endpoint');
    });

    it('should use default "MS_" prefix if none provided', () => {
        process.env.MS_CLIENT_ID = 'ms_client_id_val';
        const config = createAuthConfig(); // No prefix, defaults to MS_
        expect(config.clientId).toBe('ms_client_id_val');
    });

    it('should build tenant-specific endpoints when TENANT_ID is set', () => {
        delete process.env.MYAPP_AUTHORITY_HOST;
        delete process.env.MYAPP_TOKEN_ENDPOINT;
        delete process.env.MYAPP_AUTH_ENDPOINT;
        process.env.MYAPP_TENANT_ID = 'd508624f-a0b7-4fd3-9511-05b18ca02784';
        const config = createAuthConfig('MYAPP_');
        expect(config.tokenEndpoint).toBe('https://login.microsoftonline.com/d508624f-a0b7-4fd3-9511-05b18ca02784/oauth2/v2.0/token');
        expect(config.authEndpoint).toBe('https://login.microsoftonline.com/d508624f-a0b7-4fd3-9511-05b18ca02784/oauth2/v2.0/authorize');
    });
  });
});
// Adding a newline at the end of the file
