const fs = require('fs').promises;
const https = require('https');
const path = require('path');
const querystring = require('querystring');
const config = require('../../config');
const TokenStorage = require('../../auth/token-storage');

jest.mock('fs', () => ({
  promises: {
    readFile: jest.fn(),
    writeFile: jest.fn(),
    unlink: jest.fn(),
  },
}));
jest.mock('https');

const mockHomeDir = '/mock/home';
process.env.HOME = mockHomeDir; // Mock HOME for token path

const baseConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  redirectUri: 'http://localhost/callback',
  scopes: ['test_scope'],
  tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
};

describe('TokenStorage', () => {
  let tokenStorage;
  const tokenStorePath = path.join(mockHomeDir, '.outlook-mcp-tokens.json');

  beforeEach(() => {
    jest.resetAllMocks();
    tokenStorage = new TokenStorage(baseConfig);
    // Ensure tokens are null at the start of each test that doesn't mock readFile
    tokenStorage.tokens = null;
    tokenStorage._loadPromise = null;
    tokenStorage._refreshPromise = null;
  });

  describe('Constructor', () => {
    it('should initialize with default and provided config', () => {
      expect(tokenStorage.config.clientId).toBe('test-client-id');
      expect(tokenStorage.config.tokenStorePath).toBe(tokenStorePath);
      expect(tokenStorage.config.refreshTokenBuffer).toBe(5 * 60 * 1000);
    });

    it('should warn if client ID or secret is missing', () => {
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();
      new TokenStorage({ ...baseConfig, clientId: null });
      expect(consoleWarnSpy).toHaveBeenCalledWith(
        'TokenStorage: MS_CLIENT_ID/MS_CLIENT_SECRET (or OUTLOOK_CLIENT_ID/OUTLOOK_CLIENT_SECRET) is not configured. Token operations might fail.'
      );
      consoleWarnSpy.mockRestore();
    });

    it('should use config.js scopes as default when no MS_SCOPES env var is set', () => {
      delete process.env.MS_SCOPES;
      const storage = new TokenStorage({
        clientId: 'test-client-id',
        clientSecret: 'test-client-secret',
        redirectUri: 'http://localhost/callback',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      });
      expect(storage.config.scopes).toEqual(config.AUTH_CONFIG.scopes);
    });

    it('should use MS_SCOPES env var override when set', () => {
      process.env.MS_SCOPES = 'offline_access User.Read Mail.Read';
      const storage = new TokenStorage({
        clientId: 'test-client-id',
        clientSecret: 'test-client-secret',
        redirectUri: 'http://localhost/callback',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      });
      expect(storage.config.scopes).toEqual(['offline_access', 'User.Read', 'Mail.Read']);
      delete process.env.MS_SCOPES;
    });

    it('should warn if MS_SCOPES override is missing offline_access', () => {
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();
      process.env.MS_SCOPES = 'User.Read Mail.Read';
      new TokenStorage({
        clientId: 'test-client-id',
        clientSecret: 'test-client-secret',
        redirectUri: 'http://localhost/callback',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      });
      expect(consoleWarnSpy).toHaveBeenCalledWith(expect.stringContaining('offline_access'));
      delete process.env.MS_SCOPES;
      consoleWarnSpy.mockRestore();
    });

    it('should default flowScope from config.js FLOW_SCOPE and allow override', () => {
      expect(tokenStorage.config.flowScope).toBe(config.FLOW_SCOPE);
      const customStorage = new TokenStorage({ ...baseConfig, flowScope: 'custom-flow-scope' });
      expect(customStorage.config.flowScope).toBe('custom-flow-scope');
    });
  });

  describe('_loadTokensFromFile', () => {
    it('should load and parse tokens if file exists', async () => {
      const mockTokens = { access_token: 'loaded_token', expires_at: Date.now() + 3600000 };
      fs.readFile.mockResolvedValue(JSON.stringify(mockTokens));
      const loaded = await tokenStorage._loadTokensFromFile();
      expect(fs.readFile).toHaveBeenCalledWith(tokenStorePath, 'utf8');
      expect(loaded).toEqual(mockTokens);
      expect(tokenStorage.tokens).toEqual(mockTokens);
    });

    it('should return null and log if file does not exist (ENOENT)', async () => {
      const consoleLogSpy = jest.spyOn(console, 'log').mockImplementation();
      fs.readFile.mockRejectedValue({ code: 'ENOENT' });
      const loaded = await tokenStorage._loadTokensFromFile();
      expect(loaded).toBeNull();
      expect(tokenStorage.tokens).toBeNull();
      expect(consoleLogSpy).toHaveBeenCalledWith('Token file not found. No tokens loaded.');
      consoleLogSpy.mockRestore();
    });

    it('should return null and log error for other read errors', async () => {
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
      fs.readFile.mockRejectedValue(new Error('Read error'));
      const loaded = await tokenStorage._loadTokensFromFile();
      expect(loaded).toBeNull();
      expect(tokenStorage.tokens).toBeNull();
      expect(consoleErrorSpy).toHaveBeenCalledWith('Error loading token cache:', expect.any(Error));
      consoleErrorSpy.mockRestore();
    });
  });

  describe('_saveTokensToFile', () => {
    it('should write tokens to file', async () => {
      tokenStorage.tokens = { access_token: 'save_token' };
      await tokenStorage._saveTokensToFile();
      expect(fs.writeFile).toHaveBeenCalledWith(
        tokenStorePath,
        JSON.stringify(tokenStorage.tokens, null, 2),
        { mode: 0o600 }
      );
    });

    it('should log warning if no tokens to save', async () => {
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();
      tokenStorage.tokens = null;
      // Since it now returns false (or throws if we change it more), we expect it to return false
      const result = await tokenStorage._saveTokensToFile();
      expect(result).toBe(false); // As per current _saveTokensToFile when tokens are null
      expect(fs.writeFile).not.toHaveBeenCalled();
      expect(consoleWarnSpy).toHaveBeenCalledWith('No tokens to save.');
      consoleWarnSpy.mockRestore();
    });

    it('should throw error if fs.writeFile fails', async () => {
      tokenStorage.tokens = { access_token: 'test_token' };
      const writeError = new Error('Disk full');
      fs.writeFile.mockRejectedValue(writeError);
      await expect(tokenStorage._saveTokensToFile()).rejects.toThrow(writeError);
    });
  });

  describe('getTokens', () => {
    it('should return cached tokens if available', async () => {
      tokenStorage.tokens = { access_token: 'cached_token' };
      const tokens = await tokenStorage.getTokens();
      expect(tokens).toEqual({ access_token: 'cached_token' });
      expect(fs.readFile).not.toHaveBeenCalled();
    });

    it('should load tokens from file if not cached', async () => {
      const mockFileTokens = { access_token: 'file_token' };
      fs.readFile.mockResolvedValue(JSON.stringify(mockFileTokens));
      const tokens = await tokenStorage.getTokens();
      expect(tokens).toEqual(mockFileTokens);
      expect(fs.readFile).toHaveBeenCalledTimes(1);
    });

    it('should only call _loadTokensFromFile once for concurrent calls', async () => {
      const mockFileTokens = { access_token: 'concurrent_load_token' };
      fs.readFile.mockResolvedValue(JSON.stringify(mockFileTokens));

      const promise1 = tokenStorage.getTokens();
      const promise2 = tokenStorage.getTokens();

      const [tokens1, tokens2] = await Promise.all([promise1, promise2]);

      expect(tokens1).toEqual(mockFileTokens);
      expect(tokens2).toEqual(mockFileTokens);
      expect(fs.readFile).toHaveBeenCalledTimes(1); // Crucial check
    });
  });

  describe('getExpiryTime', () => {
    it('should return expires_at if tokens exist', () => {
      const expiry = Date.now() + 1000;
      tokenStorage.tokens = { expires_at: expiry };
      expect(tokenStorage.getExpiryTime()).toBe(expiry);
    });
    it('should return 0 if no tokens or expires_at', () => {
      tokenStorage.tokens = null;
      expect(tokenStorage.getExpiryTime()).toBe(0);
      tokenStorage.tokens = { access_token: 'no_expiry' };
      expect(tokenStorage.getExpiryTime()).toBe(0);
    });
  });

  describe('isTokenExpired', () => {
    it('should return true if no tokens or expires_at', () => {
      tokenStorage.tokens = null;
      expect(tokenStorage.isTokenExpired()).toBe(true);
      tokenStorage.tokens = { access_token: 'no_expiry_token' };
      expect(tokenStorage.isTokenExpired()).toBe(true);
    });

    it('should return true if token is past expiration time (considering buffer)', () => {
      tokenStorage.tokens = {
        expires_at: Date.now() - (tokenStorage.config.refreshTokenBuffer + 1000),
      }; // Expired by 1s + buffer
      expect(tokenStorage.isTokenExpired()).toBe(true);
    });

    it('should return true if token is within buffer period', () => {
      tokenStorage.tokens = {
        expires_at: Date.now() + (tokenStorage.config.refreshTokenBuffer - 1000),
      }; // Expires in buffer - 1s
      expect(tokenStorage.isTokenExpired()).toBe(true);
    });

    it('should return false if token is not expired and outside buffer', () => {
      tokenStorage.tokens = {
        expires_at: Date.now() + (tokenStorage.config.refreshTokenBuffer + 60000),
      }; // Valid for 1 min + buffer
      expect(tokenStorage.isTokenExpired()).toBe(false);
    });
  });

  describe('exchangeCodeForTokens', () => {
    let mockHttpsRequest;
    const mockAuthCode = 'auth_code_123';

    beforeEach(() => {
      mockHttpsRequest = {
        on: jest.fn((event, cb) => {
          if (event === 'error') mockHttpsRequest.errorHandler = cb;
          return mockHttpsRequest;
        }),
        setTimeout: jest.fn(),
        write: jest.fn(),
        end: jest.fn(),
      };
      https.request.mockImplementation((url, options, callback) => {
        mockHttpsRequest.callback = callback; // Store the callback for triggering
        return mockHttpsRequest;
      });
    });

    const mockSuccessfulTokenResponse = {
      access_token: 'new_access_token',
      refresh_token: 'new_refresh_token',
      expires_in: 3600,
      scope: 'test_scope',
      token_type: 'Bearer',
    };

    it('should successfully exchange code for tokens and save them', async () => {
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');

      // Start the exchange process
      const exchangePromise = tokenStorage.exchangeCodeForTokens(mockAuthCode);

      // Simulate successful HTTPS response
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulTokenResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes); // Trigger the https.request callback

      const tokens = await exchangePromise;

      expect(https.request).toHaveBeenCalledTimes(1);
      const requestArgs = https.request.mock.calls[0];
      expect(requestArgs[0]).toBe(baseConfig.tokenEndpoint); // URL
      expect(requestArgs[1].method).toBe('POST'); // options

      const requestBody = querystring.parse(mockHttpsRequest.write.mock.calls[0][0]);
      expect(requestBody.grant_type).toBe('authorization_code');
      expect(requestBody.code).toBe(mockAuthCode);
      expect(requestBody.client_id).toBe(baseConfig.clientId);
      expect(requestBody.scope).toBe(baseConfig.scopes.join(' '));

      expect(tokens.access_token).toBe('new_access_token');
      expect(tokenStorage.tokens.access_token).toBe('new_access_token');
      expect(tokenStorage.tokens.expires_at).toBeGreaterThan(Date.now());
      expect(saveSpy).toHaveBeenCalled();
    });

    it('should reject if saving exchanged token fails', async () => {
      const saveError = new Error('Disk space full');
      // Mock _saveTokensToFile to throw this error
      jest.spyOn(tokenStorage, '_saveTokensToFile').mockRejectedValueOnce(saveError);

      const exchangePromise = tokenStorage.exchangeCodeForTokens(mockAuthCode);
      const mockRes = {
        // Simulate successful API response
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulTokenResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(exchangePromise).rejects.toThrow(
        `Tokens exchanged but failed to save: ${saveError.message}`
      );
      // Check that tokens were updated in memory before save attempt
      expect(tokenStorage.tokens.access_token).toBe(mockSuccessfulTokenResponse.access_token);
    });

    it('should reject on token exchange API error', async () => {
      const errorResponse = { error: 'invalid_grant', error_description: 'Bad auth code' };
      const exchangePromise = tokenStorage.exchangeCodeForTokens(mockAuthCode);
      const mockRes = {
        statusCode: 400,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(errorResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(exchangePromise).rejects.toThrow(errorResponse.error_description);
    });

    it('should reject on network error during token exchange', async () => {
      const networkError = new Error('Network fail');
      const exchangePromise = tokenStorage.exchangeCodeForTokens(mockAuthCode);

      // Simulate network error by calling the 'error' handler on the request object
      mockHttpsRequest.errorHandler(networkError);

      await expect(exchangePromise).rejects.toThrow('Network fail');
    });

    it('should reject if client ID or secret is missing', async () => {
      tokenStorage.config.clientId = null;
      await expect(tokenStorage.exchangeCodeForTokens(mockAuthCode)).rejects.toThrow(
        'Client ID or Client Secret is not configured. Cannot exchange code for tokens.'
      );
    });
  });

  describe('refreshAccessToken', () => {
    let mockHttpsRequest;

    beforeEach(() => {
      mockHttpsRequest = {
        on: jest.fn((event, cb) => {
          if (event === 'error') mockHttpsRequest.errorHandler = cb;
          return mockHttpsRequest;
        }),
        setTimeout: jest.fn(),
        write: jest.fn(),
        end: jest.fn(),
      };
      https.request.mockImplementation((url, options, callback) => {
        mockHttpsRequest.callback = callback;
        return mockHttpsRequest;
      });
      tokenStorage.tokens = {
        access_token: 'old_access_token',
        refresh_token: 'valid_refresh_token',
        expires_at: Date.now() - 7200000, // Expired 2 hours ago
      };
    });

    const mockSuccessfulRefreshResponse = {
      access_token: 'refreshed_access_token',
      // refresh_token: 'optional_new_refresh_token', // MS sometimes doesn't return new one
      expires_in: 3600,
    };

    it('should successfully refresh token and save', async () => {
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');
      const refreshPromise = tokenStorage.refreshAccessToken();

      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulRefreshResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      const accessToken = await refreshPromise;
      expect(accessToken).toBe('refreshed_access_token');
      expect(tokenStorage.tokens.access_token).toBe('refreshed_access_token');
      expect(tokenStorage.tokens.expires_at).toBeGreaterThan(Date.now());
      expect(saveSpy).toHaveBeenCalled();

      const requestBody = querystring.parse(mockHttpsRequest.write.mock.calls[0][0]);
      expect(requestBody.grant_type).toBe('refresh_token');
      expect(requestBody.refresh_token).toBe('valid_refresh_token');
      expect(requestBody.scope).toBe(baseConfig.scopes.join(' '));
    });

    it('should reject if saving refreshed token fails', async () => {
      const saveError = new Error('Failed to save disk');
      // Mock _saveTokensToFile to throw an error *after* a successful API response
      jest.spyOn(tokenStorage, '_saveTokensToFile').mockRejectedValueOnce(saveError);

      const refreshPromise = tokenStorage.refreshAccessToken();
      const mockRes = {
        // Simulate successful API response
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulRefreshResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(refreshPromise).rejects.toThrow(
        `Access token refreshed but failed to save: ${saveError.message}`
      );
      // Ensure tokens in memory are updated despite save failure, as per current logic before rejection
      expect(tokenStorage.tokens.access_token).toBe(mockSuccessfulRefreshResponse.access_token);
    });

    it('should use existing refresh_token if new one is not in response', async () => {
      const refreshPromise = tokenStorage.refreshAccessToken();
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          // Response without a new refresh_token
          if (event === 'data')
            cb(
              Buffer.from(
                JSON.stringify({ ...mockSuccessfulRefreshResponse, refresh_token: undefined })
              )
            );
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);
      await refreshPromise;
      expect(tokenStorage.tokens.refresh_token).toBe('valid_refresh_token'); // Should remain the same
    });

    it('should update refresh_token if a new one is in response', async () => {
      const refreshPromise = tokenStorage.refreshAccessToken();
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data')
            cb(
              Buffer.from(
                JSON.stringify({
                  ...mockSuccessfulRefreshResponse,
                  refresh_token: 'new_returned_refresh_token',
                })
              )
            );
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);
      await refreshPromise;
      expect(tokenStorage.tokens.refresh_token).toBe('new_returned_refresh_token');
    });

    it('should reject and clear promise on refresh API error', async () => {
      const errorResponse = { error: 'invalid_grant', error_description: 'Refresh token expired' };
      const refreshPromise = tokenStorage.refreshAccessToken();
      const mockRes = {
        statusCode: 400,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(errorResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(refreshPromise).rejects.toThrow(errorResponse.error_description);
      expect(tokenStorage._refreshPromise).toBeNull();
    });

    it('should throw if no refresh token is available', async () => {
      tokenStorage.tokens.refresh_token = null;
      await expect(tokenStorage.refreshAccessToken()).rejects.toThrow(
        'No refresh token available to refresh the access token.'
      );
    });

    it('should handle concurrent refresh calls by returning the same promise', async () => {
      const promise1 = tokenStorage.refreshAccessToken();
      const promise2 = tokenStorage.refreshAccessToken();

      // Both calls should share the same underlying _refreshPromise (deduplication)
      expect(tokenStorage._refreshPromise).not.toBeNull();

      // Simulate successful response for the single underlying HTTP request
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulRefreshResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      const [accessToken1, accessToken2] = await Promise.all([promise1, promise2]);
      expect(accessToken1).toBe('refreshed_access_token');
      expect(accessToken2).toBe('refreshed_access_token');
      expect(https.request).toHaveBeenCalledTimes(1); // Crucial: only one actual HTTP request
      expect(tokenStorage._refreshPromise).toBeNull(); // Promise should be cleared after resolution
    });
  });

  describe('refreshFlowAccessToken', () => {
    let mockHttpsRequest;
    const mockFlowRefreshToken = 'valid-flow-refresh';

    beforeEach(() => {
      mockHttpsRequest = {
        on: jest.fn((event, cb) => {
          if (event === 'error') mockHttpsRequest.errorHandler = cb;
          return mockHttpsRequest;
        }),
        setTimeout: jest.fn(),
        write: jest.fn(),
        end: jest.fn(),
      };
      https.request.mockImplementation((url, options, callback) => {
        mockHttpsRequest.callback = callback;
        return mockHttpsRequest;
      });
      tokenStorage.tokens = {
        access_token: 'graph_access_token',
        refresh_token: 'graph_refresh_token',
        expires_at: Date.now() + 3600000,
        flow_access_token: 'old_flow_access_token',
        flow_refresh_token: mockFlowRefreshToken,
        flow_expires_at: Date.now() - 7200000, // Expired 2 hours ago
      };
    });

    const mockSuccessfulFlowRefreshResponse = {
      access_token: 'refreshed_flow_access_token',
      expires_in: 3600,
    };

    it('should successfully refresh flow token and save', async () => {
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');
      const refreshPromise = tokenStorage.refreshFlowAccessToken();

      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulFlowRefreshResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      const accessToken = await refreshPromise;
      expect(accessToken).toBe('refreshed_flow_access_token');
      expect(tokenStorage.tokens.flow_access_token).toBe('refreshed_flow_access_token');
      expect(tokenStorage.tokens.flow_expires_at).toBeGreaterThan(Date.now());
      expect(tokenStorage.tokens.access_token).toBe('graph_access_token');
      expect(saveSpy).toHaveBeenCalled();

      const requestBody = querystring.parse(mockHttpsRequest.write.mock.calls[0][0]);
      expect(requestBody.grant_type).toBe('refresh_token');
      expect(requestBody.refresh_token).toBe(mockFlowRefreshToken);
      expect(requestBody.scope).toBe(config.FLOW_SCOPE);
    });

    it('should update flow_refresh_token if a new one is in response', async () => {
      const refreshPromise = tokenStorage.refreshFlowAccessToken();
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data')
            cb(
              Buffer.from(
                JSON.stringify({
                  ...mockSuccessfulFlowRefreshResponse,
                  refresh_token: 'new_flow_refresh_token',
                })
              )
            );
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);
      await refreshPromise;
      expect(tokenStorage.tokens.flow_refresh_token).toBe('new_flow_refresh_token');
    });

    it('should keep existing flow_refresh_token if response does not include one', async () => {
      const refreshPromise = tokenStorage.refreshFlowAccessToken();
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data')
            cb(
              Buffer.from(
                JSON.stringify({
                  ...mockSuccessfulFlowRefreshResponse,
                  refresh_token: undefined,
                })
              )
            );
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);
      await refreshPromise;
      expect(tokenStorage.tokens.flow_refresh_token).toBe(mockFlowRefreshToken);
    });

    it('should reject and clear flow tokens on API error while preserving Graph tokens', async () => {
      const errorResponse = {
        error: 'invalid_grant',
        error_description: 'Flow refresh token expired',
      };
      const refreshPromise = tokenStorage.refreshFlowAccessToken();
      const mockRes = {
        statusCode: 400,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(errorResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(refreshPromise).rejects.toThrow(errorResponse.error_description);
      expect(tokenStorage.tokens.flow_access_token).toBeNull();
      expect(tokenStorage.tokens.flow_refresh_token).toBeNull();
      expect(tokenStorage.tokens.access_token).toBe('graph_access_token');
      expect(tokenStorage.tokens.refresh_token).toBe('graph_refresh_token');
      expect(fs.writeFile).toHaveBeenCalled();
    });

    it('should throw if no flow refresh token is available', async () => {
      tokenStorage.tokens.flow_refresh_token = null;
      await expect(tokenStorage.refreshFlowAccessToken()).rejects.toThrow(
        'No flow refresh token available to refresh the flow access token.'
      );
    });

    it('should handle concurrent refresh calls by returning the same promise', async () => {
      const promise1 = tokenStorage.refreshFlowAccessToken();
      const promise2 = tokenStorage.refreshFlowAccessToken();

      expect(tokenStorage._flowRefreshPromise).not.toBeNull();

      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(mockSuccessfulFlowRefreshResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      const [accessToken1, accessToken2] = await Promise.all([promise1, promise2]);
      expect(accessToken1).toBe('refreshed_flow_access_token');
      expect(accessToken2).toBe('refreshed_flow_access_token');
      expect(https.request).toHaveBeenCalledTimes(1);
      expect(tokenStorage._flowRefreshPromise).toBeNull();
    });
  });

  describe('getValidAccessToken', () => {
    beforeEach(() => {
      // Ensure tokens are loaded for these tests, or mock _loadTokensFromFile / getTokens
      jest.spyOn(tokenStorage, 'getTokens').mockImplementation(async () => tokenStorage.tokens);
    });

    it('should return existing token if valid and not near expiry', async () => {
      tokenStorage.tokens = { access_token: 'valid_token', expires_at: Date.now() + 3600000 }; // Expires in 1 hour
      const token = await tokenStorage.getValidAccessToken();
      expect(token).toBe('valid_token');
      expect(https.request).not.toHaveBeenCalled(); // No refresh attempt
    });

    it('should attempt refresh if token is expired and refresh token exists', async () => {
      tokenStorage.tokens = {
        access_token: 'expired_token',
        refresh_token: 'can_refresh',
        expires_at: Date.now() - 1000,
      };
      const refreshSpy = jest
        .spyOn(tokenStorage, 'refreshAccessToken')
        .mockResolvedValue('refreshed_token_from_spy');

      const token = await tokenStorage.getValidAccessToken();
      expect(refreshSpy).toHaveBeenCalled();
      expect(token).toBe('refreshed_token_from_spy');
    });

    it('should return null and clear tokens if refresh fails', async () => {
      tokenStorage.tokens = {
        access_token: 'expired_token_will_fail',
        refresh_token: 'will_fail_refresh',
        expires_at: Date.now() - 1000,
      };
      jest.spyOn(tokenStorage, 'refreshAccessToken').mockRejectedValue(new Error('Refresh failed'));
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');

      const token = await tokenStorage.getValidAccessToken();
      expect(token).toBeNull();
      expect(tokenStorage.tokens).toBeNull(); // Tokens should be invalidated
      expect(saveSpy).toHaveBeenCalled(); // Invalidation should be persisted
    });

    it('should propagate error if saving nulled token fails after refresh failure', async () => {
      tokenStorage.tokens = {
        access_token: 'expired_token_save_fail',
        refresh_token: 'refresh_me',
        expires_at: Date.now() - 1000,
      };
      jest
        .spyOn(tokenStorage, 'refreshAccessToken')
        .mockRejectedValue(new Error('Refresh API down'));
      const saveError = new Error('Disk write error during null save');
      jest.spyOn(tokenStorage, '_saveTokensToFile').mockRejectedValueOnce(saveError); // This is key

      await expect(tokenStorage.getValidAccessToken()).rejects.toThrow(saveError);
      expect(tokenStorage.tokens).toBeNull(); // Still nulled in memory
    });

    it('should return null and clear tokens if expired and no refresh token', async () => {
      tokenStorage.tokens = {
        access_token: 'expired_no_refresh',
        expires_at: Date.now() - 1000,
        // No refresh_token
      };
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile').mockResolvedValue(true); // Assume save works for this path
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();

      const token = await tokenStorage.getValidAccessToken();
      expect(token).toBeNull();
      expect(consoleWarnSpy).toHaveBeenCalledWith(
        'No refresh token available. Cannot refresh access token.'
      );
      expect(tokenStorage.tokens).toBeNull();
      expect(saveSpy).toHaveBeenCalled();
      consoleWarnSpy.mockRestore();
    });

    it('should propagate error if saving nulled token fails (no refresh token path)', async () => {
      tokenStorage.tokens = {
        access_token: 'expired_no_refresh_save_fail',
        expires_at: Date.now() - 1000,
      };
      const saveError = new Error('Disk write error during null save (no-refresh path)');
      jest.spyOn(tokenStorage, '_saveTokensToFile').mockRejectedValueOnce(saveError);

      await expect(tokenStorage.getValidAccessToken()).rejects.toThrow(saveError);
      expect(tokenStorage.tokens).toBeNull(); // Still nulled in memory
    });

    it('should return null if no tokens are loaded initially', async () => {
      tokenStorage.tokens = null; // Simulate no tokens loaded
      // Ensure getTokens returns the current null state for this specific test
      tokenStorage.getTokens.mockResolvedValue(null);

      const token = await tokenStorage.getValidAccessToken();
      expect(token).toBeNull();
    });
  });

  describe('isFlowTokenExpired', () => {
    it('should return true when no flow token or expiry is set', () => {
      tokenStorage.tokens = null;
      expect(tokenStorage.isFlowTokenExpired()).toBe(true);
      tokenStorage.tokens = { access_token: 'graph_token' };
      expect(tokenStorage.isFlowTokenExpired()).toBe(true);
    });

    it('should return false when flow token expiry is beyond buffer', () => {
      tokenStorage.tokens = {
        flow_access_token: 'flow_token',
        flow_expires_at: Date.now() + tokenStorage.config.refreshTokenBuffer + 60000,
      };
      expect(tokenStorage.isFlowTokenExpired()).toBe(false);
    });

    it('should return true when flow token is within buffer or past expiry', () => {
      tokenStorage.tokens = {
        flow_access_token: 'flow_token',
        flow_expires_at: Date.now() + tokenStorage.config.refreshTokenBuffer - 1000,
      };
      expect(tokenStorage.isFlowTokenExpired()).toBe(true);

      tokenStorage.tokens = {
        flow_access_token: 'flow_token',
        flow_expires_at: Date.now() - 60000,
      };
      expect(tokenStorage.isFlowTokenExpired()).toBe(true);
    });
  });

  describe('getFlowAccessToken', () => {
    it('should return the flow access token when valid', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'flow-token-123',
        flow_expires_at: Date.now() + 3600000,
      };
      const token = await tokenStorage.getFlowAccessToken();
      expect(token).toBe('flow-token-123');
    });

    it('should return null when flow token is expired or missing', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'expired-flow-token',
        flow_expires_at: Date.now() - 60000,
      };
      expect(await tokenStorage.getFlowAccessToken()).toBeNull();

      tokenStorage.tokens = { access_token: 'graph_token' };
      expect(await tokenStorage.getFlowAccessToken()).toBeNull();
    });
  });

  describe('saveFlowTokens', () => {
    it('should merge flow tokens with existing Graph tokens and persist', async () => {
      tokenStorage.tokens = {
        access_token: 'graph-token',
        refresh_token: 'graph-refresh',
        expires_at: Date.now() + 3600000,
      };

      await tokenStorage.saveFlowTokens({
        access_token: 'flow-token',
        refresh_token: 'flow-refresh',
        expires_in: 3600,
      });

      expect(fs.writeFile).toHaveBeenCalledWith(
        tokenStorePath,
        JSON.stringify(tokenStorage.tokens, null, 2),
        { mode: 0o600 }
      );
      expect(tokenStorage.tokens.access_token).toBe('graph-token');
      expect(tokenStorage.tokens.flow_access_token).toBe('flow-token');
      expect(tokenStorage.tokens.flow_refresh_token).toBe('flow-refresh');
      expect(tokenStorage.tokens.flow_expires_at).toBeGreaterThanOrEqual(Date.now());
    });

    it('should create a new token file when only flow tokens are provided', async () => {
      tokenStorage.tokens = {};

      await tokenStorage.saveFlowTokens({
        access_token: 'flow-token',
        refresh_token: 'flow-refresh',
        expires_at: Date.now() + 3600000,
      });

      expect(fs.writeFile).toHaveBeenCalledWith(
        tokenStorePath,
        JSON.stringify(tokenStorage.tokens, null, 2),
        { mode: 0o600 }
      );
      expect(tokenStorage.tokens.flow_access_token).toBe('flow-token');
    });
  });

  describe('getValidFlowAccessToken', () => {
    beforeEach(() => {
      jest.spyOn(tokenStorage, 'getTokens').mockImplementation(async () => tokenStorage.tokens);
    });

    it('should return the flow token when valid', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'valid-flow-token',
        flow_expires_at: Date.now() + 3600000,
      };
      const token = await tokenStorage.getValidFlowAccessToken();
      expect(token).toBe('valid-flow-token');
      expect(https.request).not.toHaveBeenCalled();
    });

    it('should refresh expired flow token when flow_refresh_token exists', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'expired-flow-token',
        flow_refresh_token: 'valid-flow-refresh',
        flow_expires_at: Date.now() - 60000,
      };
      jest
        .spyOn(tokenStorage, 'refreshFlowAccessToken')
        .mockResolvedValue('refreshed-flow-token-from-spy');

      const token = await tokenStorage.getValidFlowAccessToken();
      expect(tokenStorage.refreshFlowAccessToken).toHaveBeenCalled();
      expect(token).toBe('refreshed-flow-token-from-spy');
    });

    it('should return null and not call OAuth when flow token is expired and no flow_refresh_token exists', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'expired-flow-token',
        flow_expires_at: Date.now() - 60000,
      };
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();
      const refreshSpy = jest
        .spyOn(tokenStorage, 'refreshFlowAccessToken')
        .mockRejectedValue(new Error('Should not be called'));

      const token = await tokenStorage.getValidFlowAccessToken();
      expect(token).toBeNull();
      expect(refreshSpy).not.toHaveBeenCalled();
      expect(https.request).not.toHaveBeenCalled();
      expect(consoleWarnSpy).toHaveBeenCalledWith(
        'No flow refresh token available. Cannot refresh flow access token.'
      );
      consoleWarnSpy.mockRestore();
    });

    it('should return null and invalidate flow tokens when refresh fails with invalid_grant', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'expired-flow-token',
        flow_refresh_token: 'will-fail-flow-refresh',
        flow_expires_at: Date.now() - 60000,
        access_token: 'graph-token',
        refresh_token: 'graph-refresh',
      };
      jest
        .spyOn(tokenStorage, 'refreshFlowAccessToken')
        .mockRejectedValue(new Error('invalid_grant: Flow refresh token expired'));
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');

      const token = await tokenStorage.getValidFlowAccessToken();
      expect(token).toBeNull();
      expect(tokenStorage.tokens.flow_access_token).toBeNull();
      expect(tokenStorage.tokens.flow_refresh_token).toBeNull();
      expect(tokenStorage.tokens.access_token).toBe('graph-token');
      expect(saveSpy).toHaveBeenCalled();
    });

    it('should return null but NOT invalidate flow tokens on transient refresh failure', async () => {
      tokenStorage.tokens = {
        flow_access_token: 'expired-flow-token',
        flow_refresh_token: 'will-fail-flow-refresh',
        flow_expires_at: Date.now() - 60000,
        access_token: 'graph-token',
        refresh_token: 'graph-refresh',
      };
      jest
        .spyOn(tokenStorage, 'refreshFlowAccessToken')
        .mockRejectedValue(new Error('Flow refresh failed (transient)'));
      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile');

      const token = await tokenStorage.getValidFlowAccessToken();
      expect(token).toBeNull();
      expect(tokenStorage.tokens.flow_access_token).toBe('expired-flow-token');
      expect(tokenStorage.tokens.flow_refresh_token).toBe('will-fail-flow-refresh');
      expect(tokenStorage.tokens.access_token).toBe('graph-token');
      expect(saveSpy).not.toHaveBeenCalled();
    });

    it('should load tokens from file when not cached and return the token if valid', async () => {
      const mockFileTokens = {
        flow_access_token: 'file-flow-token',
        flow_expires_at: Date.now() + 3600000,
      };
      fs.readFile.mockResolvedValue(JSON.stringify(mockFileTokens));
      tokenStorage.tokens = null;
      tokenStorage.getTokens.mockRestore();

      const token = await tokenStorage.getValidFlowAccessToken();
      expect(fs.readFile).toHaveBeenCalledTimes(1);
      expect(token).toBe('file-flow-token');
    });
  });

  describe('clearTokens', () => {
    it('should set tokens to null and attempt to delete file', async () => {
      tokenStorage.tokens = { access_token: 'some_token' };
      fs.unlink.mockResolvedValue(); // Simulate successful deletion

      await tokenStorage.clearTokens();

      expect(tokenStorage.tokens).toBeNull();
      expect(fs.unlink).toHaveBeenCalledWith(tokenStorePath);
    });

    it('should log if token file does not exist during unlink', async () => {
      fs.unlink.mockRejectedValue({ code: 'ENOENT' });
      const consoleLogSpy = jest.spyOn(console, 'log').mockImplementation();

      await tokenStorage.clearTokens();

      expect(consoleLogSpy).toHaveBeenCalledWith('Token file not found, nothing to delete.');
      consoleLogSpy.mockRestore();
    });

    it('should log error for other unlink errors', async () => {
      fs.unlink.mockRejectedValue(new Error('Deletion failed'));
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();

      await tokenStorage.clearTokens();

      expect(consoleErrorSpy).toHaveBeenCalledWith('Error deleting token file:', expect.any(Error));
      consoleErrorSpy.mockRestore();
    });
  });
});
// Adding a newline at the end of the file
