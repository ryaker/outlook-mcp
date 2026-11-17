const handleSearchEmails = require('../../email/search');
const { WHOLE_MAILBOX_ENDPOINT } = require('../../email/search');
const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath, FOLDER_NOT_FOUND } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

describe('handleSearchEmails', () => {
  const mockAccessToken = 'dummy_access_token';
  const mockEmail = {
    id: 'email-1',
    subject: 'Invoice 12345',
    from: { emailAddress: { name: 'Example Supplier', address: 'no-reply@example.com' } },
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

    const result = await handleSearchEmails({ query: '12345' });

    expect(resolveFolderPath).not.toHaveBeenCalled();
    expect(callGraphAPIPaginated.mock.calls[0][2]).toBe(WHOLE_MAILBOX_ENDPOINT);
    expect(result.content[0].text).toContain('Found 1 emails');
    expect(result.content[0].text).toContain('Invoice 12345');
  });

  test('resolves a named folder strictly, including subfolders', async () => {
    resolveFolderPath.mockResolvedValue('me/mailFolders/abc/messages');
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail] });

    await handleSearchEmails({ query: '12345', folder: 'Suppliers' });

    expect(resolveFolderPath).toHaveBeenCalledWith(mockAccessToken, 'Suppliers', { strict: true });
    expect(callGraphAPIPaginated.mock.calls[0][2]).toBe('me/mailFolders/abc/messages');
  });

  test('reports an unknown folder instead of searching the inbox', async () => {
    resolveFolderPath.mockRejectedValue(
      Object.assign(new Error('Mail folder "Nope" not found'), { code: FOLDER_NOT_FOUND })
    );

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
  test('reports a failed folder lookup separately from a missing folder', async () => {
    resolveFolderPath.mockRejectedValue(new Error('Graph API error 503'));

    const result = await handleSearchEmails({ query: 'x', folder: 'Projects' });

    expect(callGraphAPIPaginated).not.toHaveBeenCalled();
    expect(result.content[0].text).toContain('Could not look up folder "Projects": Graph API error 503');
  });

  test('reports an error when every search request fails', async () => {
    callGraphAPIPaginated.mockRejectedValue(new Error('Graph API error 503'));

    const result = await handleSearchEmails({ query: 'invoice' });

    expect(result.content[0].text).toContain('Error searching emails: Search failed: Graph API error 503');
    expect(result.content[0].text).not.toContain('No emails found');
  });

  test('reports "no emails found" when a request succeeds after another failed', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('Graph API error 400'))
      .mockResolvedValue({ value: [] });

    const result = await handleSearchEmails({ query: 'invoice' });

    expect(result.content[0].text).toBe('No emails found matching your search criteria.');
  });
});
