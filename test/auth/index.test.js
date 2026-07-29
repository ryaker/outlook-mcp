const authIndex = require('../../auth/index');
const TokenStorage = require('../../auth/token-storage');

describe('auth/index', () => {
  it('should export tokenStorage singleton that is a TokenStorage instance', () => {
    expect(authIndex.tokenStorage).toBeDefined();
    expect(authIndex.tokenStorage).toBeInstanceOf(TokenStorage);
  });
});
