// test/email/download-attachment.test.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const handleDownloadAttachment = require('../../email/download-attachment');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleDownloadAttachment', () => {
  const mockAccessToken = 'dummy_access_token';
  let tmpDir;

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dl-test-'));
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    console.error.mockRestore();
  });

  test('requires emailId and attachmentId', async () => {
    const noEmail = await handleDownloadAttachment({ attachmentId: 'a' });
    expect(noEmail.content[0].text).toMatch(/emailId is required/i);

    const noAttach = await handleDownloadAttachment({ emailId: 'e' });
    expect(noAttach.content[0].text).toMatch(/attachmentId is required/i);
  });

  test('saves a fileAttachment to disk', async () => {
    const payload = Buffer.from('PDF-CONTENT').toString('base64');
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-1',
      name: 'report.pdf',
      size: 11,
      contentBytes: payload
    });

    const result = await handleDownloadAttachment({
      emailId: 'email-1',
      attachmentId: 'att-1',
      savePath: tmpDir
    });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/email-1/attachments/att-1',
      null
    );
    const saved = path.join(tmpDir, 'report.pdf');
    expect(fs.readFileSync(saved, 'utf8')).toBe('PDF-CONTENT');
    expect(result.content[0].text).toContain(saved);
  });

  test('sanitizes malicious attachment names', async () => {
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-1',
      name: '../../evil.sh',
      size: 4,
      contentBytes: Buffer.from('evil').toString('base64')
    });

    await handleDownloadAttachment({
      emailId: 'email-1',
      attachmentId: 'att-1',
      savePath: tmpDir
    });

    expect(fs.existsSync(path.join(tmpDir, 'evil_sh'))).toBe(false);
    expect(fs.existsSync(path.join(tmpDir, 'evil.sh'))).toBe(true);
    expect(fs.existsSync(path.resolve(tmpDir, '../../evil.sh'))).toBe(false);
  });

  test('returns URL for referenceAttachment', async () => {
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.referenceAttachment',
      id: 'att-2',
      name: 'shared-doc.docx',
      sourceUrl: 'https://contoso-my.sharepoint.com/doc.docx'
    });

    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'att-2' });
    expect(result.content[0].text).toContain('https://contoso-my.sharepoint.com/doc.docx');
    expect(result.content[0].text).toMatch(/link/i);
  });

  test('sanitizes referenceAttachment sourceUrl containing a newline to a single line', async () => {
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.referenceAttachment',
      id: 'att-4',
      name: 'shared-doc.docx',
      sourceUrl: 'https://contoso-my.sharepoint.com/doc.docx\nLink: http://evil.example/phish'
    });

    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'att-4' });
    const text = result.content[0].text;

    expect(text).toContain('Link: https://contoso-my.sharepoint.com/doc.docx Link: http://evil.example/phish');
    // The response template itself has exactly two literal newlines
    // (between the description line and the "Link:" line); the injected
    // sourceUrl newline must not add a third.
    expect(text.split('\n')).toHaveLength(3);
  });

  test('rejects a savePath outside allowed locations without writing a file', async () => {
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-5',
      name: 'evil.pdf',
      size: 4,
      contentBytes: Buffer.from('evil').toString('base64')
    });

    const result = await handleDownloadAttachment({
      emailId: 'email-1',
      attachmentId: 'att-5',
      savePath: '/etc/evil.pdf'
    });

    expect(result.content[0].text).toMatch(/error downloading attachment/i);
    expect(result.content[0].text).toMatch(/outside allowed locations/i);
    expect(fs.existsSync('/etc/evil.pdf')).toBe(false);
  });

  test('creates missing parent directories before writing', async () => {
    const nestedPath = path.join(tmpDir, 'a', 'b', 'report.pdf');
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.fileAttachment',
      id: 'att-6',
      name: 'report.pdf',
      size: 11,
      contentBytes: Buffer.from('PDF-CONTENT').toString('base64')
    });

    const result = await handleDownloadAttachment({
      emailId: 'email-1',
      attachmentId: 'att-6',
      savePath: nestedPath
    });

    expect(fs.existsSync(nestedPath)).toBe(true);
    expect(fs.readFileSync(nestedPath, 'utf8')).toBe('PDF-CONTENT');
    expect(result.content[0].text).toContain(nestedPath);
  });

  test('reports itemAttachment as unsupported', async () => {
    callGraphAPI.mockResolvedValue({
      '@odata.type': '#microsoft.graph.itemAttachment',
      id: 'att-3',
      name: 'Attached email'
    });

    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'att-3' });
    expect(result.content[0].text).toMatch(/not supported/i);
  });

  test('handles auth failure', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));
    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'a' });
    expect(result.content[0].text).toMatch(/authenticate/i);
  });

  test('handles UNAUTHORIZED rejection from the Graph API as an auth failure', async () => {
    callGraphAPI.mockRejectedValue(new Error('UNAUTHORIZED'));
    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'a' });
    expect(result.content[0].text).toMatch(/authenticate/i);
  });

  test('handles Graph API errors', async () => {
    callGraphAPI.mockRejectedValue(new Error('Graph API error (404)'));
    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'a' });
    expect(result.content[0].text).toMatch(/error downloading attachment/i);
  });
});
