const fs = require('fs').promises;
const https = require('https');
const querystring = require('querystring');
const TokenStorage = require('../../auth/token-storage');

jest.mock('fs', () => ({
  promises: {
    readFile: jest.fn(),
    writeFile: jest.fn(),
    unlink: jest.fn(),
  }
}));
jest.mock('https');

const mockHomeDir = '/mock/home';
process.env.HOME = mockHomeDir;

const baseConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  redirectUri: 'http://localhost/callback',
  scopes: ['offline_access', 'Mail.Read'],
  tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
};

describe('Device Code Flow', () => {
  let tokenStorage;
  let mockHttpsRequest;

  beforeEach(() => {
    jest.resetAllMocks();
    tokenStorage = new TokenStorage(baseConfig);
    tokenStorage.tokens = null;
    tokenStorage._loadPromise = null;
    tokenStorage._refreshPromise = null;

    mockHttpsRequest = {
      on: jest.fn((event, cb) => {
        if (event === 'error') mockHttpsRequest.errorHandler = cb;
        return mockHttpsRequest;
      }),
      write: jest.fn(),
      end: jest.fn(),
    };
    https.request.mockImplementation((url, options, callback) => {
      mockHttpsRequest.callback = callback;
      return mockHttpsRequest;
    });
  });

  describe('initiateDeviceCodeFlow', () => {
    it('should call the devicecode endpoint and return the response', async () => {
      const deviceCodeResponse = {
        device_code: 'test-device-code',
        user_code: 'ABCD1234',
        verification_uri: 'https://microsoft.com/devicelogin',
        interval: 5,
        expires_in: 900,
        message: 'To sign in, go to https://microsoft.com/devicelogin and enter the code ABCD1234',
      };

      const promise = tokenStorage.initiateDeviceCodeFlow();

      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(deviceCodeResponse)));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      const result = await promise;

      expect(result.user_code).toBe('ABCD1234');
      expect(result.device_code).toBe('test-device-code');
      expect(result.verification_uri).toBe('https://microsoft.com/devicelogin');

      // Verify correct endpoint was called (devicecode, not token)
      const calledUrl = https.request.mock.calls[0][0];
      expect(calledUrl).toContain('/devicecode');

      // Verify request body contains client_id and scope
      const requestBody = querystring.parse(mockHttpsRequest.write.mock.calls[0][0]);
      expect(requestBody.client_id).toBe('test-client-id');
      expect(requestBody.scope).toContain('offline_access');
    });

    it('should reject on error response', async () => {
      const promise = tokenStorage.initiateDeviceCodeFlow();

      const mockRes = {
        statusCode: 400,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify({
            error: 'invalid_client',
            error_description: 'Client is not supported for device code flow',
          })));
          if (event === 'end') cb();
        },
      };
      mockHttpsRequest.callback(mockRes);

      await expect(promise).rejects.toThrow('Client is not supported for device code flow');
    });
  });

  describe('pollForDeviceCodeToken', () => {
    beforeEach(() => {
      jest.useFakeTimers();
    });

    afterEach(() => {
      jest.useRealTimers();
    });

    it('should poll and return tokens on success', async () => {
      const tokenResponse = {
        access_token: 'device-access-token',
        refresh_token: 'device-refresh-token',
        expires_in: 3600,
        scope: 'Mail.Read',
        token_type: 'Bearer',
      };

      const saveSpy = jest.spyOn(tokenStorage, '_saveTokensToFile').mockResolvedValue();

      // First poll: authorization_pending, second poll: success
      let callCount = 0;
      https.request.mockImplementation((url, options, callback) => {
        const req = {
          on: jest.fn().mockReturnThis(),
          write: jest.fn(),
          end: jest.fn(() => {
            callCount++;
            const mockRes = {
              statusCode: callCount === 1 ? 400 : 200,
              on: (event, cb) => {
                if (event === 'data') {
                  if (callCount === 1) {
                    cb(Buffer.from(JSON.stringify({
                      error: 'authorization_pending',
                      error_description: 'User has not yet authenticated',
                    })));
                  } else {
                    cb(Buffer.from(JSON.stringify(tokenResponse)));
                  }
                }
                if (event === 'end') cb();
              },
            };
            callback(mockRes);
          }),
        };
        return req;
      });

      const pollPromise = tokenStorage.pollForDeviceCodeToken('test-device-code', 5, 900);
      // Advance past the poll interval to trigger the second request
      await jest.advanceTimersByTimeAsync(6000);

      const result = await pollPromise;

      expect(result.access_token).toBe('device-access-token');
      expect(result.refresh_token).toBe('device-refresh-token');
      expect(saveSpy).toHaveBeenCalled();
      expect(callCount).toBe(2);
    });

    it('should reject on authorization_declined', async () => {
      https.request.mockImplementation((url, options, callback) => {
        const req = {
          on: jest.fn().mockReturnThis(),
          write: jest.fn(),
          end: jest.fn(() => {
            const mockRes = {
              statusCode: 400,
              on: (event, cb) => {
                if (event === 'data') cb(Buffer.from(JSON.stringify({
                  error: 'authorization_declined',
                  error_description: 'User declined the authorization',
                })));
                if (event === 'end') cb();
              },
            };
            callback(mockRes);
          }),
        };
        return req;
      });

      await expect(
        tokenStorage.pollForDeviceCodeToken('test-device-code', 5, 900)
      ).rejects.toThrow('User declined the authorization');
    });

    it('should reject on timeout', async () => {
      // Use real timers for this test — set expiresIn to 0 so it times out immediately
      jest.useRealTimers();

      https.request.mockImplementation((url, options, callback) => {
        const req = {
          on: jest.fn().mockReturnThis(),
          write: jest.fn(),
          end: jest.fn(() => {
            const mockRes = {
              statusCode: 400,
              on: (event, cb) => {
                if (event === 'data') cb(Buffer.from(JSON.stringify({
                  error: 'authorization_pending',
                })));
                if (event === 'end') cb();
              },
            };
            callback(mockRes);
          }),
        };
        return req;
      });

      // expiresIn=0 means deadline is already passed after first poll
      await expect(
        tokenStorage.pollForDeviceCodeToken('test-device-code', 0, 0)
      ).rejects.toThrow('timed out');
    });
  });
});
