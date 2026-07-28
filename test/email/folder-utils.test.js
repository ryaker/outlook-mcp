const {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  resolveSegmentInParent
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

  describe('getFolderIdByName - path resolution', () => {
    beforeEach(() => {
      callGraphAPI.mockReset();
    });

    test('two-level path resolves to subfolder ID', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [{ id: 'tramite-id', displayName: 'Tramite' }] })
        .mockResolvedValueOnce({ value: [{ id: 'req-104951-id', displayName: 'REQ-104951' }] });

      const result = await getFolderIdByName(mockAccessToken, 'Tramite/REQ-104951');

      expect(result).toBe('req-104951-id');
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        1,
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: "displayName eq 'Tramite'" }
      );
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        2,
        mockAccessToken,
        'GET',
        'me/mailFolders/tramite-id/childFolders',
        null,
        { $filter: "displayName eq 'REQ-104951'" }
      );
    });

    test('three-level path resolves to nested subfolder ID', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [{ id: 'a-id', displayName: 'A' }] })
        .mockResolvedValueOnce({ value: [{ id: 'b-id', displayName: 'B' }] })
        .mockResolvedValueOnce({ value: [{ id: 'c-id', displayName: 'C' }] });

      const result = await getFolderIdByName(mockAccessToken, 'A/B/C');

      expect(result).toBe('c-id');
      expect(callGraphAPI).toHaveBeenCalledTimes(3);
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        3,
        mockAccessToken,
        'GET',
        'me/mailFolders/b-id/childFolders',
        null,
        { $filter: "displayName eq 'C'" }
      );
    });

    test('flat folder name uses single top-level API call', async () => {
      callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'inbox-id', displayName: 'Inbox' }] });

      const result = await getFolderIdByName(mockAccessToken, 'Inbox');

      expect(result).toBe('inbox-id');
      expect(callGraphAPI).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: "displayName eq 'Inbox'" }
      );
    });

    test('case-insensitive segments resolve via fallback', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [{ id: 'tramite-id', displayName: 'TRAMITE' }] })
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [{ id: 'req-104951-id', displayName: 'req-104951' }] });

      const result = await getFolderIdByName(mockAccessToken, 'tramite/req-104951');

      expect(result).toBe('req-104951-id');
      expect(callGraphAPI).toHaveBeenCalledTimes(4);
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        2,
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $top: 100 }
      );
      expect(callGraphAPI).toHaveBeenNthCalledWith(
        4,
        mockAccessToken,
        'GET',
        'me/mailFolders/tramite-id/childFolders',
        null,
        { $top: 100 }
      );
    });

    test('non-existent segment returns null and short-circuits', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [{ id: 'tramite-id', displayName: 'Tramite' }] })
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [{ id: 'other-id', displayName: 'OtherFolder' }] });

      const result = await getFolderIdByName(mockAccessToken, 'Tramite/NONEXISTENT/Child');

      expect(result).toBeNull();
      expect(callGraphAPI).toHaveBeenCalledTimes(3);
    });

    test('non-existent top-level folder returns null with no child queries', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [{ id: 'other-id', displayName: 'OtherFolder' }] });

      const result = await getFolderIdByName(mockAccessToken, 'NonExistent/Child');

      expect(result).toBeNull();
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('empty segments are filtered and path still resolves', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [{ id: 'tramite-id', displayName: 'Tramite' }] })
        .mockResolvedValueOnce({ value: [{ id: 'req-104951-id', displayName: 'REQ-104951' }] });

      const result = await getFolderIdByName(mockAccessToken, 'Tramite//REQ-104951');

      expect(result).toBe('req-104951-id');
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('all-empty path returns null without API calls', async () => {
      const result1 = await getFolderIdByName(mockAccessToken, '/');
      const result2 = await getFolderIdByName(mockAccessToken, '//');

      expect(result1).toBeNull();
      expect(result2).toBeNull();
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('API error mid-traversal returns null', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [{ id: 'a-id', displayName: 'A' }] })
        .mockRejectedValueOnce(new Error('API Error'));

      const result = await getFolderIdByName(mockAccessToken, 'A/B');

      expect(result).toBeNull();
      expect(callGraphAPI).toHaveBeenCalledTimes(2);
    });

    test('resolveSegmentInParent exact match at top-level', async () => {
      callGraphAPI.mockResolvedValueOnce({ value: [{ id: 'top-id', displayName: 'TopFolder' }] });

      const result = await resolveSegmentInParent(mockAccessToken, null, 'TopFolder');

      expect(result).toBe('top-id');
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders',
        null,
        { $filter: "displayName eq 'TopFolder'" }
      );
    });

    test('resolveSegmentInParent case-insensitive fallback at child level', async () => {
      callGraphAPI
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({ value: [{ id: 'child-id', displayName: 'childname' }] });

      const result = await resolveSegmentInParent(mockAccessToken, 'parent-id', 'CHILDNAME');

      expect(result).toBe('child-id');
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders/parent-id/childFolders',
        null,
        { $top: 100 }
      );
    });
  });
});
