const { handleCheckAuthStatus } = require('../../auth/tools');
const TokenStorage = require('../../auth/token-storage');

jest.mock('../../auth/token-storage');

describe('auth/tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('handleCheckAuthStatus', () => {
    it('returns "Authenticated and ready" when getValidAccessToken returns a token', async () => {
      const mockInstance = {
        getValidAccessToken: jest.fn().mockResolvedValue('valid_access_token')
      };
      TokenStorage.mockImplementation(() => mockInstance);

      const result = await handleCheckAuthStatus();

      expect(TokenStorage).toHaveBeenCalledTimes(1);
      expect(mockInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        content: [{ type: 'text', text: 'Authenticated and ready' }]
      });
    });

    it('returns "Not authenticated" when getValidAccessToken returns null', async () => {
      const mockInstance = {
        getValidAccessToken: jest.fn().mockResolvedValue(null)
      };
      TokenStorage.mockImplementation(() => mockInstance);

      const result = await handleCheckAuthStatus();

      expect(TokenStorage).toHaveBeenCalledTimes(1);
      expect(mockInstance.getValidAccessToken).toHaveBeenCalledTimes(1);
      expect(result).toEqual({
        content: [{ type: 'text', text: 'Not authenticated' }]
      });
    });
  });
});
// Adding a newline at the end of the file
