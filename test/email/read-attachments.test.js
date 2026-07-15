const handleReadEmail = require('../../email/read');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('read-email attachment listing', () => {
  const mockAccessToken = 'dummy_access_token';

  const baseEmail = {
    id: 'email-1',
    subject: 'Test',
    from: { emailAddress: { name: 'John', address: 'john@example.com' } },
    toRecipients: [{ emailAddress: { name: 'Me', address: 'me@example.com' } }],
    receivedDateTime: '2024-01-15T10:30:00Z',
    body: { contentType: 'text', content: 'Hello' },
    hasAttachments: true
  };

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('lists attachments when hasAttachments is true', async () => {
    callGraphAPI
      .mockResolvedValueOnce(baseEmail)
      .mockResolvedValueOnce({
        value: [
          { id: 'att-1', name: 'report.pdf', size: 251392, contentType: 'application/pdf', isInline: false },
          { id: 'att-2', name: 'logo.png', size: 12288, contentType: 'image/png', isInline: true }
        ]
      });

    const result = await handleReadEmail({ id: 'email-1' });
    const text = result.content[0].text;

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-1/attachments',
      null,
      { $select: 'id,name,size,contentType,isInline' }
    );
    expect(text).toContain('Attachments (2):');
    expect(text).toContain('report.pdf (245.5 KB) — application/pdf [id: att-1]');
    expect(text).toContain('logo.png (12 KB, inline) — image/png [id: att-2]');
  });

  test('skips attachment call when hasAttachments is false', async () => {
    callGraphAPI.mockResolvedValueOnce({ ...baseEmail, hasAttachments: false });

    const result = await handleReadEmail({ id: 'email-1' });

    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(result.content[0].text).not.toContain('Attachments (');
  });

  test('email still returned when attachment listing fails', async () => {
    callGraphAPI
      .mockResolvedValueOnce(baseEmail)
      .mockRejectedValueOnce(new Error('Graph API error (500)'));

    const result = await handleReadEmail({ id: 'email-1' });
    const text = result.content[0].text;

    expect(text).toContain('Subject: Test');
    expect(text).toContain('Could not list attachments');
  });
});
