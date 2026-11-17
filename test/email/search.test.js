const handleSearchEmails = require('../../email/search');
const { WHOLE_MAILBOX_ENDPOINT } = require('../../email/search');
const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

describe('handleSearchEmails', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockEmail = {
    id: 'email-1',
    subject: 'Factuur 20104260',
    from: { emailAddress: { name: 'Techno Lease', address: 'no-reply@example.com' } },
    receivedDateTime: '2026-10-02T17:05:37Z',
    isRead: true
  };

  beforeEach(() => {
    jest.resetAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('searches the whole mailbox when no folder is given', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail] });

    const result = await handleSearchEmails({ query: '20104260' });

    expect(resolveFolderPath).not.toHaveBeenCalled();
    expect(callGraphAPIPaginated.mock.calls[0][2]).toBe(WHOLE_MAILBOX_ENDPOINT);
    expect(result.content[0].text).toContain('Found 1 emails');
    expect(result.content[0].text).toContain('Factuur 20104260');
  });

  test('resolves a named folder strictly, including subfolders', async () => {
    resolveFolderPath.mockResolvedValue('me/mailFolders/abc/messages');
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail] });

    await handleSearchEmails({ query: '20104260', folder: 'Suppliers WP' });

    expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, 'Suppliers WP', { strict: true });
    expect(callGraphAPIPaginated.mock.calls[0][2]).toBe('me/mailFolders/abc/messages');
  });

  test('reports an unknown folder instead of searching the inbox', async () => {
    resolveFolderPath.mockRejectedValue(new Error('Mail folder "Nope" not found'));

    const result = await handleSearchEmails({ query: 'x', folder: 'Nope' });

    expect(callGraphAPIPaginated).not.toHaveBeenCalled();
    expect(result.content[0].text).toContain('Folder "Nope" not found');
  });

  test('returns "no emails found" instead of recent emails when nothing matches', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [] });

    const result = await handleSearchEmails({ query: 'does-not-exist', hasAttachments: true });

    expect(result.content[0].text).toBe('No emails found matching your search criteria.');
    // No call may run without a $search term (that would return unrelated emails)
    for (const call of callGraphAPIPaginated.mock.calls) {
      expect(call[3].$search).toBeDefined();
    }
  });

  test('still supports filter-only searches without search terms', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail] });

    const result = await handleSearchEmails({ unreadOnly: true });

    expect(callGraphAPIPaginated.mock.calls[0][3].$filter).toBe('isRead eq false');
    expect(result.content[0].text).toContain('Found 1 emails');
  });
});
