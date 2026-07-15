/**
 * Helpers for saving email attachments to disk safely.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

/**
 * Sanitize an attachment filename so it cannot escape the target directory.
 * Strips bidirectional-override/isolate characters (used to visually spoof
 * file extensions, e.g. an RLO character making "evil.exe" display as
 * "evil.pdf"), null bytes, converts path separators to underscores, removes
 * leading dots. Falls back to 'attachment' for degenerate names.
 * @param {string} name - Untrusted filename from the attachment
 * @returns {string} - Safe filename
 */
function sanitizeFilename(name) {
  if (!name || typeof name !== 'string') {
    return 'attachment';
  }
  let safe = name
    .replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g, '')
    .replace(/\u0000/g, '')
    .replace(/^[/\\]+/, '')
    .replace(/\.\.[/\\]/g, '')
    .replace(/[/\\]/g, '_')
    .replace(/^\.+/, '')
    .trim();
  return safe || 'attachment';
}

/**
 * Resolve the final absolute path to save an attachment to.
 * - undefined savePath -> ~/Downloads/<filename>
 * - savePath is an existing directory -> <savePath>/<filename>
 * - otherwise savePath is the full target file path
 * A leading ~ is expanded to the home directory. If the resolved path
 * already exists, a " (N)" suffix is added before the extension.
 *
 * savePath is model-controlled input: it comes from the tool call arguments,
 * and prompt injection via email content could try to steer it toward
 * sensitive files (e.g. ~/.ssh/authorized_keys, /etc/passwd). The resolved
 * target is therefore restricted to the user's home directory or a temp
 * directory, and hidden files/folders (any path segment starting with '.')
 * are blocked within those roots too.
 * @param {string|undefined} savePath - User-supplied directory or file path
 * @param {string} filename - Sanitized attachment filename
 * @returns {string} - Absolute path that does not currently exist
 * @throws {Error} - If the resolved target is outside the allowed roots or
 *   targets a hidden file/folder
 */
function resolveSavePath(savePath, filename) {
  let target;
  if (!savePath) {
    target = path.join(os.homedir(), 'Downloads', filename);
  } else {
    let expanded = savePath;
    if (expanded === '~' || expanded.startsWith('~/')) {
      expanded = path.join(os.homedir(), expanded.slice(1));
    }
    expanded = path.resolve(expanded);
    if (fs.existsSync(expanded) && fs.statSync(expanded).isDirectory()) {
      target = path.join(expanded, filename);
    } else {
      target = expanded;
    }
  }
  if (!isPathAllowed(target)) {
    throw new Error('savePath outside allowed locations (home or temp directory) or targets a hidden file/folder.');
  }
  return dedupePath(target);
}

/**
 * Check whether target is inside one of the allowed roots (home directory
 * or a temp directory), and that no path segment relative to that root
 * starts with '.' (which would target a hidden file/folder, e.g. ~/.ssh).
 * @param {string} target - Absolute path to check
 * @returns {boolean} - True if target is allowed
 */
function isPathAllowed(target) {
  const allowedRoots = [os.homedir(), os.tmpdir(), '/tmp', '/private/tmp'];
  for (const root of allowedRoots) {
    const rel = path.relative(root, target);
    const isInside = rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
    if (!isInside) {
      continue;
    }
    const segments = rel === '' ? [] : rel.split(path.sep);
    const hasHiddenSegment = segments.some((segment) => segment.startsWith('.'));
    if (!hasHiddenSegment) {
      return true;
    }
  }
  return false;
}

/**
 * If target exists, append " (1)", " (2)", ... before the extension.
 */
function dedupePath(target) {
  if (!fs.existsSync(target)) {
    return target;
  }
  const dir = path.dirname(target);
  const ext = path.extname(target);
  const base = path.basename(target, ext);
  let counter = 1;
  let candidate;
  do {
    candidate = path.join(dir, `${base} (${counter})${ext}`);
    counter++;
  } while (fs.existsSync(candidate));
  return candidate;
}

/**
 * Sanitize an untrusted display string (e.g. an attachment name or content
 * type) for safe single-line rendering in tool output text. Attachment
 * metadata comes from the email sender and is untrusted; without this,
 * newlines/control characters could forge fake list entries or tool output
 * headers. Strips bidirectional-override/isolate characters (which can
 * visually reorder or hide parts of the rendered text), replaces C0 control
 * characters (0x00-0x1F, including \r \n \t) and DEL (0x7F) with a space,
 * collapses repeated spaces to one, and trims. Printable characters
 * (including hyphens) pass through unchanged.
 * This is for display only — it does not make a value safe to use as a
 * filesystem path; use sanitizeFilename for that.
 * @param {*} value - Untrusted display value (e.g. attachment name/contentType)
 * @returns {string} - Single-line safe string; '' for null/undefined
 */
function sanitizeDisplayName(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value)
    .replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g, '')
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/ +/g, ' ')
    .trim();
}

/**
 * Format a byte count as a human-readable string.
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = { sanitizeFilename, resolveSavePath, formatSize, sanitizeDisplayName };
