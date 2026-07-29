const fs = require('fs').promises;

jest.mock('fs', () => ({
  promises: {
    readFile: jest.fn(),
    writeFile: jest.fn(),
    unlink: jest.fn(),
  },
}));

const mockHomeDir = '/mock/home';
process.env.HOME = mockHomeDir;

const handlerModules = [
  '../../power-automate/list-environments',
  '../../power-automate/list-flows',
  '../../power-automate/list-runs',
  '../../power-automate/run-flow',
  '../../power-automate/toggle-flow',
];

const requiredArgs = {
  'list-flows': { environmentId: 'Default-12345' },
  'list-runs': { environmentId: 'Default-12345', flowId: 'flow-123' },
  'run-flow': { environmentId: 'Default-12345', flowId: 'flow-123' },
  'toggle-flow': { environmentId: 'Default-12345', flowId: 'flow-123' },
};

describe('power-automate handlers', () => {
  let tokenStorage;

  beforeEach(() => {
    jest.resetModules();
    jest.resetAllMocks();
    fs.readFile.mockRejectedValue({ code: 'ENOENT' });
    tokenStorage = require('../../auth/index').tokenStorage;
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  handlerModules.forEach((handlerPath) => {
    const name = handlerPath.replace('../../power-automate/', '');

    it(`${name} imports tokenStorage and returns auth required when no flow token`, async () => {
      jest.spyOn(tokenStorage, 'getValidFlowAccessToken').mockResolvedValue(null);

      const handler = require(handlerPath);
      const args = requiredArgs[name] || {};
      const result = await handler(args);

      expect(tokenStorage.getValidFlowAccessToken).toHaveBeenCalled();
      expect(result.content[0].text).toMatch(/Power Automate authentication required/);
    });
  });
});
