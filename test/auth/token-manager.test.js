const fs = require('fs');
const tokenManager = require('../../auth/token-manager');

jest.mock('fs');

describe('token-manager', () => {
  beforeEach(() => {
    jest.resetAllMocks();
  });

  it('should still create test tokens', () => {
    fs.existsSync.mockReturnValue(false);
    fs.writeFileSync.mockImplementation(() => {});

    const tokens = tokenManager.createTestTokens();

    expect(tokens).toHaveProperty('access_token');
    expect(tokens).toHaveProperty('refresh_token');
    expect(tokens).toHaveProperty('expires_at');
    expect(fs.writeFileSync).toHaveBeenCalled();
  });

  it('should not export getFlowAccessToken or saveFlowTokens', () => {
    expect(tokenManager.getFlowAccessToken).toBeUndefined();
    expect(tokenManager.saveFlowTokens).toBeUndefined();
    expect(tokenManager.createTestTokens).toBeDefined();
  });
});
