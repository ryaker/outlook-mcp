const fs = require('fs');
const os = require('os');
const path = require('path');
const { sanitizeFilename, resolveSavePath, formatSize } = require('../../email/attachment-utils');

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
});

describe('formatSize', () => {
  test('formats bytes, KB, MB', () => {
    expect(formatSize(0)).toBe('0 B');
    expect(formatSize(500)).toBe('500 B');
    expect(formatSize(251392)).toBe('245.5 KB');
    expect(formatSize(1048576)).toBe('1 MB');
  });
});
