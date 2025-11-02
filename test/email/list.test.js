const handleListEmails = require('../../email/list');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath, WELL_KNOWN_FOLDERS } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

describe('handleListEmails', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockEmails = [
    {
      id: 'email-1',
      subject: 'Test Email 1',
      from: {
        emailAddress: {
          name: 'John Doe',
          address: 'john@example.com'
        }
      },
      receivedDateTime: '2024-01-15T10:30:00Z',
      isRead: false
    },
    {
      id: 'email-2',
      subject: 'Test Email 2',
      from: {
        emailAddress: {
          name: 'Jane Smith',
          address: 'jane@example.com'
        }
      },
      receivedDateTime: '2024-01-14T15:20:00Z',
      isRead: true
    }
  ];

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    resolveFolderPath.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful email retrieval', () => {
    test('should list emails from inbox by default', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: mockEmails });

      const result = await handleListEmails({});

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, 'inbox');
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        WELL_KNOWN_FOLDERS['inbox'],
        null,
        expect.objectContaining({
          $top: 10,
          $orderby: 'receivedDateTime desc'
        })
      );
      expect(result.content[0].text).toContain('Found 2 emails in inbox');
      expect(result.content[0].text).toContain('Test Email 1');
      expect(result.content[0].text).toContain('[UNREAD]');
    });

    test('should list emails from specified folder', async () => {
      const customFolder = 'drafts';
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['drafts']);
      callGraphAPI.mockResolvedValue({ value: mockEmails });

      const result = await handleListEmails({ folder: customFolder });

      expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, customFolder);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        WELL_KNOWN_FOLDERS['drafts'],
        null,
        expect.any(Object)
      );
      expect(result.content[0].text).toContain('Found 2 emails in drafts');
    });

    test('should respect custom count parameter', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: [mockEmails[0]] });

      await handleListEmails({ count: 5 });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        WELL_KNOWN_FOLDERS['inbox'],
        null,
        expect.objectContaining({
          $top: 5
        })
      );
    });

    test('should format email list correctly with sender info', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: mockEmails });

      const result = await handleListEmails({});

      expect(result.content[0].text).toContain('John Doe (john@example.com)');
      expect(result.content[0].text).toContain('Jane Smith (jane@example.com)');
      expect(result.content[0].text).toContain('Subject: Test Email 1');
      expect(result.content[0].text).toContain('ID: email-1');
    });

    test('should handle email without sender info', async () => {
      const emailWithoutSender = {
        id: 'email-3',
        subject: 'No Sender Email',
        receivedDateTime: '2024-01-13T12:00:00Z',
        isRead: true
      };

      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: [emailWithoutSender] });

      const result = await handleListEmails({});

      expect(result.content[0].text).toContain('Unknown (unknown)');
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no emails found', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListEmails({});

      expect(result.content[0].text).toBe('No emails found in inbox.');
    });

    test('should return appropriate message when folder has no emails', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['archive']);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListEmails({ folder: 'archive' });

      expect(result.content[0].text).toBe('No emails found in archive.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleListEmails({});

      expect(result.content[0].text).toBe(
        "Authentication required. Please use the 'authenticate' tool first."
      );
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('should handle Graph API error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleListEmails({});

      expect(result.content[0].text).toBe('Error listing emails: Graph API Error');
    });

    test('should handle folder resolution error', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockRejectedValue(new Error('Folder resolution failed'));

      const result = await handleListEmails({ folder: 'InvalidFolder' });

      expect(result.content[0].text).toBe('Error listing emails: Folder resolution failed');
    });
  });

  describe('inbox endpoint verification', () => {
    test('should use me/mailFolders/inbox/messages for inbox folder', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      resolveFolderPath.mockResolvedValue(WELL_KNOWN_FOLDERS['inbox']);
      callGraphAPI.mockResolvedValue({ value: mockEmails });

      await handleListEmails({ folder: 'inbox' });

      expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, 'inbox');
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/mailFolders/inbox/messages',
        null,
        expect.any(Object)
      );
    });
  });
});
