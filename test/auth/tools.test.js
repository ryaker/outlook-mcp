describe('handleCheckAuthStatus', () => {
  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();
  });

  it('returns not authenticated when no token file is present', async () => {
    const getTokens = jest.fn().mockResolvedValue(null);
    const getValidAccessToken = jest.fn();

    jest.doMock('../../auth/token-storage', () => {
      return jest.fn().mockImplementation(() => ({
        getTokens,
        getValidAccessToken,
      }));
    });

    const { handleCheckAuthStatus } = require('../../auth/tools');

    const result = await handleCheckAuthStatus();

    expect(result).toEqual({
      content: [{ type: 'text', text: 'Not authenticated' }],
    });
    expect(getValidAccessToken).not.toHaveBeenCalled();
  });

  it('returns authenticated when a valid access token can be obtained', async () => {
    const getTokens = jest.fn().mockResolvedValue({
      access_token: 'cached-token',
      expires_at: Date.now() + 60_000,
    });
    const getValidAccessToken = jest.fn().mockResolvedValue('fresh-token');

    jest.doMock('../../auth/token-storage', () => {
      return jest.fn().mockImplementation(() => ({
        getTokens,
        getValidAccessToken,
      }));
    });

    const { handleCheckAuthStatus } = require('../../auth/tools');

    const result = await handleCheckAuthStatus();

    expect(getTokens).toHaveBeenCalledTimes(1);
    expect(getValidAccessToken).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      content: [{ type: 'text', text: 'Authenticated and ready' }],
    });
  });

  it('returns authenticated when the cached token is expired but refresh succeeds', async () => {
    const getTokens = jest.fn().mockResolvedValue({
      access_token: 'expired-token',
      expires_at: Date.now() - 60_000,
      refresh_token: 'refresh-token',
    });
    const getValidAccessToken = jest.fn().mockResolvedValue('refreshed-token');

    jest.doMock('../../auth/token-storage', () => {
      return jest.fn().mockImplementation(() => ({
        getTokens,
        getValidAccessToken,
      }));
    });

    const { handleCheckAuthStatus } = require('../../auth/tools');

    const result = await handleCheckAuthStatus();

    expect(getTokens).toHaveBeenCalledTimes(1);
    expect(getValidAccessToken).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      content: [{ type: 'text', text: 'Authenticated and ready' }],
    });
  });
});
