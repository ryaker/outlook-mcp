const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName
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
