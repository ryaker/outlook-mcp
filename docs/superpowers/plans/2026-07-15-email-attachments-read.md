# Email Attachment Read Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let users see attachment metadata when reading an email and download individual attachments to local disk.

**Architecture:** Pure path/filename helpers live in `email/attachment-utils.js`. A new `download-attachment` tool handler (`email/download-attachment.js`) fetches a single attachment via Graph and writes it to disk. `email/read.js` gains a second Graph call to list attachment metadata when `hasAttachments` is true. Tool wiring follows the existing `email/index.js` pattern.

**Tech Stack:** Node.js (CommonJS), Microsoft Graph REST API via `utils/graph-api.js` `callGraphAPI`, Jest for tests (existing `test/email/` conventions: `jest.mock` the graph client and auth).

**Spec:** `docs/superpowers/specs/2026-07-15-email-attachments-read-design.md`

## Global Constraints

- Read-only scope: no send/draft attachment support.
- Filenames from attachments are untrusted: strip path separators and null bytes before writing to disk (no path traversal).
- Default download directory: `~/Downloads`.
- On filename collision, save as `name (1).ext`, `name (2).ext`, …
- Error-handling pattern matches existing handlers: `Authentication required` message on auth failure; otherwise error text in the MCP response. Handlers never throw.
- All MCP responses use `{ content: [{ type: "text", text: ... }] }`.

---

### Task 1: Attachment path/filename helpers

**Files:**
- Create: `email/attachment-utils.js`
- Test: `test/email/attachment-utils.test.js`

**Interfaces:**
- Consumes: nothing project-specific (Node `fs`, `path`, `os`)
- Produces:
  - `sanitizeFilename(name: string) -> string` — strips path separators, null bytes, leading dots; falls back to `'attachment'` when empty
  - `resolveSavePath(savePath: string|undefined, filename: string) -> string` — absolute target file path; handles default dir, dir-vs-file, `~` expansion, collision rename
  - `formatSize(bytes: number) -> string` — human-readable size, e.g. `245.5 KB`

- [ ] **Step 1: Write the failing tests**

```js
// test/email/attachment-utils.test.js
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx jest test/email/attachment-utils.test.js`
Expected: FAIL — `Cannot find module '../../email/attachment-utils'`

- [ ] **Step 3: Write the implementation**

```js
// email/attachment-utils.js
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
 * Format a byte count as a human-readable string.
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = { sanitizeFilename, resolveSavePath, formatSize };
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx jest test/email/attachment-utils.test.js`
Expected: PASS (all tests)

- [ ] **Step 5: Commit**

```bash
git add email/attachment-utils.js test/email/attachment-utils.test.js
git commit -m "feat: attachment filename/path helpers with traversal protection"
```

---

### Task 2: download-attachment handler

**Files:**
- Create: `email/download-attachment.js`
- Test: `test/email/download-attachment.test.js`

**Interfaces:**
- Consumes: `sanitizeFilename`, `resolveSavePath`, `formatSize` from `email/attachment-utils.js` (Task 1); `callGraphAPI(accessToken, method, path, data, queryParams)` from `utils/graph-api.js`; `ensureAuthenticated()` from `auth/`.
- Produces: `handleDownloadAttachment(args) -> Promise<MCP response>` where `args = { emailId, attachmentId, savePath? }`. Default export (`module.exports = handleDownloadAttachment`).

- [ ] **Step 1: Write the failing tests**

```js
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

  test('handles Graph API errors', async () => {
    callGraphAPI.mockRejectedValue(new Error('Graph API error (404)'));
    const result = await handleDownloadAttachment({ emailId: 'e', attachmentId: 'a' });
    expect(result.content[0].text).toMatch(/error downloading attachment/i);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx jest test/email/download-attachment.test.js`
Expected: FAIL — `Cannot find module '../../email/download-attachment'`

- [ ] **Step 3: Write the implementation**

```js
// email/download-attachment.js
/**
 * Download email attachment functionality
 *
 * Security: attachment filenames are untrusted input. They are sanitized
 * before writing to disk to prevent path traversal.
 */
const fs = require('fs');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeFilename, resolveSavePath, formatSize } = require('./attachment-utils');

/**
 * Download attachment handler
 * @param {object} args - Tool arguments
 * @param {string} args.emailId - Email (message) ID
 * @param {string} args.attachmentId - Attachment ID from read-email listing
 * @param {string} [args.savePath] - Target directory or file path (default: ~/Downloads)
 * @returns {object} - MCP response
 */
async function handleDownloadAttachment(args) {
  const { emailId, attachmentId, savePath } = args;

  if (!emailId) {
    return {
      content: [{ type: "text", text: "emailId is required." }]
    };
  }

  if (!attachmentId) {
    return {
      content: [{ type: "text", text: "attachmentId is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(emailId)}/attachments/${encodeURIComponent(attachmentId)}`;
    const attachment = await callGraphAPI(accessToken, 'GET', endpoint, null);

    if (!attachment) {
      return {
        content: [{ type: "text", text: "Attachment not found." }]
      };
    }

    const odataType = attachment['@odata.type'];

    if (odataType === '#microsoft.graph.referenceAttachment') {
      return {
        content: [{
          type: "text",
          text: `"${attachment.name}" is a link to a cloud file, not a stored attachment.\n\nLink: ${attachment.sourceUrl || 'unavailable'}`
        }]
      };
    }

    if (odataType !== '#microsoft.graph.fileAttachment') {
      return {
        content: [{
          type: "text",
          text: `Downloading this attachment type is not supported (${odataType || 'unknown'}). It may be an attached email, contact, or calendar item — open it directly in Outlook instead.`
        }]
      };
    }

    if (!attachment.contentBytes) {
      return {
        content: [{ type: "text", text: "Attachment has no content to save." }]
      };
    }

    const filename = sanitizeFilename(attachment.name);
    const targetPath = resolveSavePath(savePath, filename);
    const buffer = Buffer.from(attachment.contentBytes, 'base64');
    fs.writeFileSync(targetPath, buffer);

    return {
      content: [{
        type: "text",
        text: `Saved: ${targetPath} (${formatSize(buffer.length)})`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    console.error(`Error downloading attachment: ${error.message}`);
    return {
      content: [{
        type: "text",
        text: `Error downloading attachment: ${error.message}`
      }]
    };
  }
}

module.exports = handleDownloadAttachment;
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx jest test/email/download-attachment.test.js`
Expected: PASS (all tests)

- [ ] **Step 5: Commit**

```bash
git add email/download-attachment.js test/email/download-attachment.test.js
git commit -m "feat: download-attachment handler"
```

---

### Task 3: Attachment listing in read-email

**Files:**
- Modify: `email/read.js` (add attachment listing after body formatting, around line 104)
- Test: `test/email/read-attachments.test.js`

**Interfaces:**
- Consumes: `formatSize` from `email/attachment-utils.js` (Task 1); existing `callGraphAPI`.
- Produces: `read-email` output gains an `Attachments (N):` section listing `name (size) — contentType [id: ...]` per attachment when `hasAttachments` is true. Attachment ids shown here are what `download-attachment` (Task 2) consumes.

- [ ] **Step 1: Write the failing tests**

```js
// test/email/read-attachments.test.js
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npx jest test/email/read-attachments.test.js`
Expected: FAIL — output lacks `Attachments (2):` (email body renders, no attachment section)

- [ ] **Step 3: Modify email/read.js**

Add import at the top (after the existing requires, `email/read.js:10`):

```js
const { formatSize } = require('./attachment-utils');
```

Add a fetch + format block after the body extraction and before `const formattedEmail = ...` (currently `email/read.js:96`):

```js
      // List attachment metadata (names/sizes/ids) when the email has attachments.
      // Failure here must not block returning the email itself.
      let attachmentSection = '';
      if (email.hasAttachments) {
        try {
          const attachResponse = await callGraphAPI(
            accessToken,
            'GET',
            `me/messages/${encodeURIComponent(emailId)}/attachments`,
            null,
            { $select: 'id,name,size,contentType,isInline' }
          );
          const attachments = (attachResponse && attachResponse.value) || [];
          if (attachments.length > 0) {
            const lines = attachments.map((a, i) => {
              const inline = a.isInline ? ', inline' : '';
              return `${i + 1}. ${a.name} (${formatSize(a.size)}${inline}) — ${a.contentType || 'unknown type'} [id: ${a.id}]`;
            });
            attachmentSection = `\n\nAttachments (${attachments.length}):\n${lines.join('\n')}\nUse the 'download-attachment' tool with an attachment id to save one to disk.`;
          }
        } catch (attachError) {
          console.error(`Error listing attachments: ${attachError.message}`);
          attachmentSection = '\n\n[Could not list attachments: ' + attachError.message + ']';
        }
      }
```

Append the section to the formatted output — change the existing return (currently `email/read.js:112-119`) from:

```js
      return {
        content: [
          {
            type: "text",
            text: formattedEmail + rawHtmlSection
          }
        ]
      };
```

to:

```js
      return {
        content: [
          {
            type: "text",
            text: formattedEmail + attachmentSection + rawHtmlSection
          }
        ]
      };
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx jest test/email/read-attachments.test.js`
Expected: PASS (all 3 tests)

- [ ] **Step 5: Run full test suite for regressions**

Run: `npx jest`
Expected: PASS (all suites, including existing `test/email/list.test.js`)

- [ ] **Step 6: Commit**

```bash
git add email/read.js test/email/read-attachments.test.js
git commit -m "feat: list attachment metadata in read-email"
```

---

### Task 4: Tool wiring + test-mode mocks

**Files:**
- Modify: `email/index.js` (import + tool definition + exports)
- Modify: `utils/mock-data.js` (attachment mock responses in `simulateGraphAPIResponse`)

**Interfaces:**
- Consumes: `handleDownloadAttachment` (Task 2, default export of `email/download-attachment.js`).
- Produces: MCP tool `download-attachment` registered in `emailTools`; test-mode responses for `GET .../attachments` and `GET .../attachments/{id}`.

- [ ] **Step 1: Wire the tool in email/index.js**

Add import after the existing handler imports (`email/index.js:10`):

```js
const handleDownloadAttachment = require('./download-attachment');
```

Add tool definition to the `emailTools` array (after the `read-email` entry, `email/index.js:94`):

```js
  {
    name: "download-attachment",
    description: "Downloads an email attachment to the local filesystem. Get attachment IDs from read-email output. Cloud-file links (reference attachments) return their URL instead.",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "ID of the email containing the attachment"
        },
        attachmentId: {
          type: "string",
          description: "ID of the attachment (shown in read-email output)"
        },
        savePath: {
          type: "string",
          description: "Where to save the file: an existing directory (attachment keeps its own name) or a full file path. Default: ~/Downloads"
        }
      },
      required: ["emailId", "attachmentId"]
    },
    handler: handleDownloadAttachment
  },
```

Add to `module.exports` (`email/index.js:215-224`):

```js
  handleDownloadAttachment,
```

- [ ] **Step 2: Add test-mode mocks in utils/mock-data.js**

Inside `simulateGraphAPIResponse`, in the `if (method === 'GET')` branch, add BEFORE the existing `if (path.includes('messages')...)` block (`utils/mock-data.js:17`) so attachment paths are matched first:

```js
    if (path.includes('/attachments/')) {
      // Single attachment download
      return {
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'simulated-attachment-1',
        name: 'simulated-report.pdf',
        contentType: 'application/pdf',
        size: 24,
        isInline: false,
        contentBytes: Buffer.from('Simulated PDF content :)').toString('base64')
      };
    }
    if (path.includes('/attachments')) {
      // Attachment listing
      return {
        value: [
          {
            id: 'simulated-attachment-1',
            name: 'simulated-report.pdf',
            contentType: 'application/pdf',
            size: 24,
            isInline: false
          }
        ]
      };
    }
```

Also set `hasAttachments: true` in the existing single-email mock (the `if (path.includes('/messages/'))` block, `utils/mock-data.js:19-40`) so the test-mode flow exercises the listing. If the mock email object has no `hasAttachments` field, add it; if it has `hasAttachments: false`, change to `true`.

- [ ] **Step 3: Verify no regressions and server loads**

Run: `npx jest`
Expected: PASS (all suites)

Run: `node -e "const { emailTools } = require('./email'); const t = emailTools.find(x => x.name === 'download-attachment'); console.log(t ? 'tool registered: ' + t.name : 'MISSING'); console.log('handler type:', typeof t.handler);"`
Expected output:
```
tool registered: download-attachment
handler type: function
```

- [ ] **Step 4: End-to-end smoke test in test mode**

Run:
```bash
node -e "
process.env.USE_TEST_MODE = 'true';
const handleReadEmail = require('./email/read');
const handleDownload = require('./email/download-attachment');
(async () => {
  const auth = require('./auth');
  // test-mode tokens start with test_access_token_
  auth.ensureAuthenticated = async () => 'test_access_token_x';
  const read = await handleReadEmail({ id: 'simulated-email-id' });
  console.log(read.content[0].text);
})();
"
```
Expected: email output includes `Attachments (1):` and `simulated-report.pdf`.

Note: if overriding `ensureAuthenticated` this way doesn't take effect (module destructuring), verify instead via `USE_TEST_MODE=true npm run inspect` and call `read-email` with id `simulated-email-id`, then `download-attachment` with `emailId: simulated-email-id`, `attachmentId: simulated-attachment-1`, `savePath: /tmp`. Expected: `Saved: /tmp/simulated-report.pdf (24 B)`.

- [ ] **Step 5: Commit**

```bash
git add email/index.js utils/mock-data.js
git commit -m "feat: register download-attachment tool, add test-mode attachment mocks"
```

---

## Verification Checklist

- [ ] `npx jest` — full suite green
- [ ] `read-email` on a real email with attachments lists names/sizes/ids (live, after `npm run auth-server` + authenticate)
- [ ] `download-attachment` saves a real PDF to `~/Downloads`, collision produces ` (1)` suffix
- [ ] `read-email` on an email without attachments is unchanged (single Graph call)
