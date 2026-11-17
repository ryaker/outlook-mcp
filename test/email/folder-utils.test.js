const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  FOLDER_NOT_FOUND
} = require('../../email/folder-utils');
const { callGraphAPI } = require('../../utils/graph-api');

jest.mock('../../utils/graph-api');

describe('resolveFolderPath', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    // Mock console.error to avoid cluttering test output
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('well-known folders', () => {
    test('should return inbox endpoint when no folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, null);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when undefined folder name is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, undefined);
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return inbox endpoint when empty string is provided', async () => {
      const result = await resolveFolderPath(mockAccessToken, '');
      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should return correct endpoint for well-known folders', async () => {
      const result = await resolveFolderPath(mockAccessToken, 'drafts');
      expect(result).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle case-insensitive well-known folder names', async () => {
      const result1 = await resolveFolderPath(mockAccessToken, 'INBOX');
      const result2 = await resolveFolderPath(mockAccessToken, 'Drafts');
      const result3 = await resolveFolderPath(mockAccessToken, 'SENT');

      expect(result1).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(result2).toBe(WELL_KNOWN_FOLDERS['drafts']);
      expect(result3).toBe(WELL_KNOWN_FOLDERS['sent']);
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });

  describe('custom folders', () => {
    test('should resolve custom folder by ID when found', async () => {
      const customFolderId = 'custom-folder-id-123';
      const customFolderName = 'MyCustomFolder';

      callGraphAPI.mockResolvedValueOnce({
        value: [{ id: customFolderId, displayName: customFolderName }]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: `displayName eq '${customFolderName}'` }
      );
    });

    test('should try case-insensitive search when exact match fails', async () => {
      const customFolderId = 'custom-folder-id-456';
      const customFolderName = 'ProjectAlpha';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns all folders for case-insensitive match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'other-id', displayName: 'OtherFolder' },
          { id: customFolderId, displayName: 'projectalpha' }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(`me/mailFolders/${customFolderId}/messages`);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should fall back to inbox when custom folder is not found', async () => {
      const nonExistentFolder = 'NonExistentFolder';

      // First call returns empty (exact match fails)
      callGraphAPI.mockResolvedValueOnce({ value: [] });

      // Second call returns folders without a match
      callGraphAPI.mockResolvedValueOnce({
        value: [
          { id: 'id1', displayName: 'Folder1' },
          { id: 'id2', displayName: 'Folder2' }
        ]
      });

      const result = await resolveFolderPath(mockAccessToken, nonExistentFolder);

      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('should fall back to inbox when API call fails', async () => {
      const customFolderName = 'CustomFolder';

      callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

      const result = await resolveFolderPath(mockAccessToken, customFolderName);

      expect(result).toBe(WELL_KNOWN_FOLDERS['inbox']);
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
    });
  });
});

describe('getFolderIdByName', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should return folder ID when exact match is found', async () => {
    const folderId = 'folder-id-123';
    const folderName = 'TestFolder';

    callGraphAPI.mockResolvedValueOnce({
      value: [{ id: folderId, displayName: folderName }]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName}'` }
    );
  });

  test('should return folder ID when case-insensitive match is found', async () => {
    const folderId = 'folder-id-456';
    const folderName = 'TestFolder';

    // First call returns empty (exact match fails)
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns folders with case-insensitive match
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: folderId, displayName: 'testfolder' }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBe(folderId);
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('should return null when folder is not found', async () => {
    const folderName = 'NonExistentFolder';

    // First call returns empty
    callGraphAPI.mockResolvedValueOnce({ value: [] });

    // Second call returns folders without a match
    callGraphAPI.mockResolvedValueOnce({
      value: [
        { id: 'id1', displayName: 'OtherFolder' }
      ]
    });

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('should return null when API call fails', async () => {
    const folderName = 'TestFolder';

    callGraphAPI.mockRejectedValueOnce(new Error('API Error'));

    const result = await getFolderIdByName(mockAccessToken, folderName);

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });
});

describe('nested folder lookup', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockReset();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('getFolderIdByName with deep finds a nested subfolder', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] }) // exact top-level match
      .mockResolvedValueOnce({ value: [{ id: 'top', displayName: 'Projects', childFolderCount: 1 }] })
      .mockResolvedValueOnce({ value: [{ id: 'nested', displayName: 'Suppliers', childFolderCount: 0 }] });

    const result = await getFolderIdByName(mockAccessToken, 'suppliers', { deep: true });

    expect(result).toBe('nested');
    expect(callGraphAPI.mock.calls[2][2]).toBe('me/mailFolders/top/childFolders');
  });

  test('getFolderIdByName without deep does not search subfolders', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [{ id: 'top', displayName: 'Projects', childFolderCount: 1 }] });

    const result = await getFolderIdByName(mockAccessToken, 'Suppliers');

    expect(result).toBeNull();
    expect(callGraphAPI).toHaveBeenCalledTimes(2);
  });

  test('resolveFolderPath resolves a nested folder', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [{ id: 'top', displayName: 'Projects', childFolderCount: 1 }] })
      .mockResolvedValueOnce({ value: [{ id: 'nested', displayName: 'Suppliers', childFolderCount: 0 }] });

    const result = await resolveFolderPath(mockAccessToken, 'Suppliers');

    expect(result).toBe('me/mailFolders/nested/messages');
  });

  test('resolveFolderPath in strict mode throws when the folder is missing', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [] });

    await expect(resolveFolderPath(mockAccessToken, 'Nope', { strict: true }))
      .rejects.toThrow('Mail folder "Nope" not found');
  });
  test('escapes apostrophes in the folder name filter', async () => {
    callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'f1', displayName: "Client's" }] });

    await getFolderIdByName(mockAccessToken, "Client's");

    expect(callGraphAPI.mock.calls[0][4].$filter).toBe("displayName eq 'Client''s'");
  });

  test('follows nextLink pages of top-level and child folders', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] }) // exact top-level match
      .mockResolvedValueOnce({
        value: [{ id: 'a', displayName: 'A', childFolderCount: 0 }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/mailFolders?$skip=100'
      })
      .mockResolvedValueOnce({ value: [{ id: 'top', displayName: 'Projects', childFolderCount: 1 }] })
      .mockResolvedValueOnce({
        value: [{ id: 'c1', displayName: 'Other', childFolderCount: 0 }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/mailFolders/top/childFolders?$skip=100'
      })
      .mockResolvedValueOnce({ value: [{ id: 'nested', displayName: 'Suppliers', childFolderCount: 0 }] });

    const result = await getFolderIdByName(mockAccessToken, 'Suppliers', { deep: true });

    expect(result).toBe('nested');
    expect(callGraphAPI.mock.calls[2][2]).toContain('$skip=100');
    expect(callGraphAPI.mock.calls[4][2]).toContain('childFolders?$skip=100');
  });

  test('strict mode propagates lookup failures instead of reporting not found', async () => {
    callGraphAPI.mockRejectedValueOnce(new Error('Graph API error 503'));

    const error = await resolveFolderPath(mockAccessToken, 'Projects', { strict: true }).catch(e => e);

    expect(error.message).toBe('Graph API error 503');
    expect(error.code).toBeUndefined();
  });

  test('strict mode propagates child folder request failures', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [{ id: 'top', displayName: 'Projects', childFolderCount: 1 }] })
      .mockRejectedValueOnce(new Error('Graph API error 429'));

    await expect(resolveFolderPath(mockAccessToken, 'Suppliers', { strict: true }))
      .rejects.toThrow('Graph API error 429');
  });

  test('strict mode marks a missing folder with FOLDER_NOT_FOUND', async () => {
    callGraphAPI
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [] });

    const error = await resolveFolderPath(mockAccessToken, 'Nope', { strict: true }).catch(e => e);

    expect(error.code).toBe(FOLDER_NOT_FOUND);
  });
  test('limits concurrent child folder requests to stay under the mailbox limit', async () => {
    const parents = Array.from({ length: 8 }, (_, i) => ({ id: `p${i}`, displayName: `P${i}`, childFolderCount: 1 }));
    let active = 0;
    let peak = 0;
    callGraphAPI.mockImplementation(async (token, method, path) => {
      if (!path.includes('childFolders')) {
        return { value: path === 'me/mailFolders' && callGraphAPI.mock.calls.length === 1 ? [] : parents };
      }
      active++;
      peak = Math.max(peak, active);
      await new Promise(resolve => setTimeout(resolve, 5));
      active--;
      return { value: [] };
    });

    const result = await getFolderIdByName(mockAccessToken, 'Missing', { deep: true });

    expect(result).toBeNull();
    expect(peak).toBeLessThanOrEqual(3);
    expect(callGraphAPI.mock.calls.filter(c => c[2].includes('childFolders'))).toHaveLength(8);
  });
});
