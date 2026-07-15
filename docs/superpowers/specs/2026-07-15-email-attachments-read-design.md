# Email Attachment Read Support — Design

**Date:** 2026-07-15
**Scope:** Read-only attachment support: list attachments on received email, download to local disk. Send-side attachments explicitly out of scope.

## Goal

Currently the server only exposes a `hasAttachments` boolean (filter in `list-emails`/`search-emails`, flag in `read-email`). Users cannot see attachment names or retrieve files. Add the ability to see attachment metadata when reading an email and download individual attachments to the local filesystem.

## Design

### 1. Extend `read-email` (`email/read.js`)

After fetching the email, if `hasAttachments` is true, make a second Graph call:

```
GET me/messages/{id}/attachments?$select=id,name,size,contentType,isInline
```

Append an attachment section to the tool output:

```
Attachments (2):
1. report.pdf (245 KB) — application/pdf [id: AAMkAG...]
2. logo.png (12 KB, inline) — image/png [id: AAMkAG...]
```

If the attachment-list call fails, the email is still returned normally with an error note in the attachment section. Existing behavior is unchanged for emails without attachments.

### 2. New tool: `download-attachment` (`email/download-attachment.js`)

**Parameters:**
- `emailId` (string, required) — the message id
- `attachmentId` (string, required) — from the `read-email` attachment listing
- `savePath` (string, optional) — defaults to `~/Downloads`. If the path is an existing directory, the attachment's own (sanitized) name is used inside it; otherwise the path is treated as the full target file path.

**Behavior:**

Fetch `GET me/messages/{emailId}/attachments/{attachmentId}` and branch on `@odata.type`:

- `#microsoft.graph.fileAttachment` — decode `contentBytes` (base64) and write the file to disk. Filename comes from the attachment `name`, sanitized: strip path separators and null bytes so a malicious attachment name cannot escape the target directory (no path traversal).
- `#microsoft.graph.referenceAttachment` — no file content exists on the message; return the OneDrive/SharePoint URL to the user.
- `#microsoft.graph.itemAttachment` (attached email/contact/event) — report unsupported and suggest opening the item directly.

Filename collision: if `report.pdf` exists at the destination, save as `report (1).pdf` (incrementing).

**Success response:** `Saved: /Users/x/Downloads/report.pdf (245 KB)`

**Error handling:** follows the existing module pattern — `Authentication required` message on auth failure, error text in the MCP response otherwise.

### 3. Wiring (`email/index.js`)

Import the handler and add the tool definition to the `EMAIL_TOOLS` array, same pattern as existing email tools.

### 4. Test mode (`utils/mock-data.js`)

Add mock responses keyed off `USE_TEST_MODE` like existing mocks:
- mock attachment list for a message
- mock `fileAttachment` with a small base64 payload so download can be exercised end-to-end

## Testing

- **Jest unit tests:** filename sanitization, collision rename, `@odata.type` branching — with `callGraphAPI` mocked.
- **Manual:** real email with a PDF attachment, both in test mode and live.

## Out of Scope

- Sending/drafting emails with attachments
- Large attachment handling beyond what a single GET returns (Graph returns full `contentBytes` for file attachments on this endpoint)
- Inline-image rendering; inline attachments are listed and downloadable like any other file
