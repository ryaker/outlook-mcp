/**
 * Helpers for saving email attachments to disk safely.
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

/**
 * Sanitize an attachment filename so it cannot escape the target directory.
 * Strips null bytes, converts path separators to underscores, removes
 * leading dots. Falls back to 'attachment' for degenerate names.
 * @param {string} name - Untrusted filename from the attachment
 * @returns {string} - Safe filename
 */
function sanitizeFilename(name) {
  if (!name || typeof name !== 'string') {
    return 'attachment';
  }
  let safe = name
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
 * @param {string|undefined} savePath - User-supplied directory or file path
 * @param {string} filename - Sanitized attachment filename
 * @returns {string} - Absolute path that does not currently exist
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
  return dedupePath(target);
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
 * headers. Replaces C0 control characters (0x00-0x1F, including \r \n \t)
 * and DEL (0x7F) with a space, collapses repeated spaces to one, and trims.
 * Printable characters (including hyphens) pass through unchanged.
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
