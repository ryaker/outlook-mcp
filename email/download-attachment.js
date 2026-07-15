/**
 * Download email attachment functionality
 *
 * Security: attachment filenames are untrusted input. They are sanitized
 * before writing to disk to prevent path traversal.
 */
const fs = require('fs');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { sanitizeFilename, resolveSavePath, formatSize, sanitizeDisplayName } = require('./attachment-utils');

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
      const displayName = sanitizeDisplayName(attachment.name);
      const displayUrl = sanitizeDisplayName(attachment.sourceUrl) || 'unavailable';
      return {
        content: [{
          type: "text",
          text: `"${displayName}" is a link to a cloud file, not a stored attachment.\n\nLink: ${displayUrl}`
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
