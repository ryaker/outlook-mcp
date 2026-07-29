const { handleCheckAuthStatus } = require('../../auth/tools');
const { tokenStorage } = require('../../auth/index');
const config = require('../../config');

jest.mock('../../auth/index', () => ({
  tokenStorage: {
    getValidAccessToken: jest.fn(),
  },
}));

describe('auth/tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('handleCheckAuthStatus', () => {
    it('returns "Authenticated and ready" when getValidAccessToken returns a token', async () => {
      tokenStorage.getValidAccessToken.mockResolvedValue('valid_access_token');

      const result = await handleCheckAuthStatus();

      expect(tokenStorage.getValidAccessToken).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        content: [{ type: 'text', text: 'Authenticated and ready' }],
      });
    });

    it('returns "Not authenticated" when getValidAccessToken returns null', async () => {
      tokenStorage.getValidAccessToken.mockResolvedValue(null);

      const result = await handleCheckAuthStatus();

      expect(tokenStorage.getValidAccessToken).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        content: [{ type: 'text', text: 'Not authenticated' }],
      });
    });
  });
});
// Adding a newline at the end of the file
