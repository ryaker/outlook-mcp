const fs = require('fs');
const os = require('os');
const path = require('path');
const { sanitizeFilename, resolveSavePath, formatSize, sanitizeDisplayName } = require('../../email/attachment-utils');

describe('sanitizeFilename', () => {
  test('passes normal filenames through', () => {
    expect(sanitizeFilename('report.pdf')).toBe('report.pdf');
  });

  test('strips path separators and traversal segments', () => {
    expect(sanitizeFilename('../../etc/passwd')).toBe('etc_passwd');
    expect(sanitizeFilename('..\\..\\win\\evil.exe')).toBe('win_evil.exe');
    expect(sanitizeFilename('/absolute/path.txt')).toBe('absolute_path.txt');
  });

  test('strips null bytes', () => {
    expect(sanitizeFilename('evil\u0000.pdf')).toBe('evil.pdf');
  });

  test('strips leading dots (no hidden files)', () => {
    expect(sanitizeFilename('...hidden')).toBe('hidden');
  });

  test('falls back when name is empty or degenerate', () => {
    expect(sanitizeFilename('')).toBe('attachment');
    expect(sanitizeFilename('...')).toBe('attachment');
    expect(sanitizeFilename(undefined)).toBe('attachment');
  });

  test('strips bidirectional override characters', () => {
    expect(sanitizeFilename('report\u202Efdp.exe')).toBe('reportfdp.exe');
  });
});

describe('resolveSavePath', () => {
  let tmpDir;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'attach-test-'));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  test('uses filename inside an existing directory', () => {
    expect(resolveSavePath(tmpDir, 'report.pdf')).toBe(path.join(tmpDir, 'report.pdf'));
  });

  test('treats non-directory path as full file path', () => {
    const target = path.join(tmpDir, 'renamed.pdf');
    expect(resolveSavePath(target, 'report.pdf')).toBe(target);
  });

  test('defaults to ~/Downloads when savePath is undefined', () => {
    const expected = path.join(os.homedir(), 'Downloads', 'report.pdf');
    expect(resolveSavePath(undefined, 'report.pdf')).toBe(expected);
  });

  test('expands leading ~', () => {
    const result = resolveSavePath('~/some-dir-that-does-not-exist.pdf', 'report.pdf');
    expect(result).toBe(path.join(os.homedir(), 'some-dir-that-does-not-exist.pdf'));
  });

  test('renames on collision: name (1).ext', () => {
    fs.writeFileSync(path.join(tmpDir, 'report.pdf'), 'x');
    expect(resolveSavePath(tmpDir, 'report.pdf')).toBe(path.join(tmpDir, 'report (1).pdf'));
  });

  test('increments past multiple collisions', () => {
    fs.writeFileSync(path.join(tmpDir, 'report.pdf'), 'x');
    fs.writeFileSync(path.join(tmpDir, 'report (1).pdf'), 'x');
    expect(resolveSavePath(tmpDir, 'report.pdf')).toBe(path.join(tmpDir, 'report (2).pdf'));
  });

  test('collision rename applies to explicit file paths too', () => {
    const target = path.join(tmpDir, 'renamed.pdf');
    fs.writeFileSync(target, 'x');
    expect(resolveSavePath(target, 'report.pdf')).toBe(path.join(tmpDir, 'renamed (1).pdf'));
  });

  test('rejects a target outside home/temp directories', () => {
    expect(() => resolveSavePath('/etc/passwd', 'report.pdf'))
      .toThrow(/outside allowed locations/);
  });

  test('rejects a hidden file/folder under the home directory', () => {
    const hidden = path.join(os.homedir(), '.ssh', 'authorized_keys');
    expect(() => resolveSavePath(hidden, 'report.pdf'))
      .toThrow(/outside allowed locations/);
  });

  test('rejects a hidden file/folder under a temp directory', () => {
    const hidden = path.join(tmpDir, '.hidden', 'report.pdf');
    expect(() => resolveSavePath(hidden, 'report.pdf'))
      .toThrow(/outside allowed locations/);
  });
});

describe('formatSize', () => {
  test('formats bytes, KB, MB', () => {
    expect(formatSize(0)).toBe('0 B');
    expect(formatSize(500)).toBe('500 B');
    expect(formatSize(251392)).toBe('245.5 KB');
    expect(formatSize(1048576)).toBe('1 MB');
  });
});

describe('sanitizeDisplayName', () => {
  test('replaces newlines and carriage returns with spaces', () => {
    expect(sanitizeDisplayName('evil\nname')).toBe('evil name');
    expect(sanitizeDisplayName('a\r\nb')).toBe('a b');
  });

  test('replaces tabs and other control characters with spaces', () => {
    expect(sanitizeDisplayName('a\tb')).toBe('a b');
    expect(sanitizeDisplayName('a\u0000b')).toBe('a b');
    expect(sanitizeDisplayName('a\u001bb')).toBe('a b'); // ESC
    expect(sanitizeDisplayName('a\u007fb')).toBe('a b'); // DEL
  });

  test('leaves hyphens and other printable characters unchanged', () => {
    expect(sanitizeDisplayName('my-report-v2.pdf')).toBe('my-report-v2.pdf');
    expect(sanitizeDisplayName('https://contoso-my.sharepoint.com/doc.docx'))
      .toBe('https://contoso-my.sharepoint.com/doc.docx');
  });

  test('collapses repeated whitespace and trims the result', () => {
    expect(sanitizeDisplayName('  a   b  ')).toBe('a b');
    expect(sanitizeDisplayName('evil\n\n\nAttachments (99):')).toBe('evil Attachments (99):');
  });

  test('returns an empty string for null or undefined', () => {
    expect(sanitizeDisplayName(null)).toBe('');
    expect(sanitizeDisplayName(undefined)).toBe('');
  });

  test('coerces non-string values to strings', () => {
    expect(sanitizeDisplayName(123)).toBe('123');
  });

  test('strips bidirectional override characters', () => {
    expect(sanitizeDisplayName('evil\u202Ename')).toBe('evilname');
  });
});
