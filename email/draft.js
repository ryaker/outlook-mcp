/**
 * Draft email functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Draft email handler
 * Creates a draft in Outlook using Microsoft Graph:
 * POST /me/messages
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDraftEmail(args) {
  const { to, cc, bcc, subject = '', body = '', importance = 'normal' } = args || {};

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Format recipients only when provided
    const toRecipients = to
      ? to.split(',').map(email => ({
          emailAddress: { address: email.trim() }
        })).filter(r => r.emailAddress.address)
      : [];

    const ccRecipients = cc
      ? cc.split(',').map(email => ({
          emailAddress: { address: email.trim() }
        })).filter(r => r.emailAddress.address)
      : [];

    const bccRecipients = bcc
      ? bcc.split(',').map(email => ({
          emailAddress: { address: email.trim() }
        })).filter(r => r.emailAddress.address)
      : [];

    // Create message payload for draft creation
    const messageObject = {
      subject,
      body: {
        contentType: typeof body === 'string' && body.toLowerCase().includes('<html') ? 'html' : 'text',
        content: body
      },
      toRecipients: toRecipients.length > 0 ? toRecipients : undefined,
      ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
      bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
      importance
    };

    // Create draft message
    const draft = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);

    return {
      content: [{
        type: "text",
        text: `Draft created successfully!\n\nDraft ID: ${draft.id}\nSubject: ${draft.subject || '(no subject)'}\nRecipients: ${toRecipients.length}${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ''}${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ''}`
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

    if (error.message && error.message.includes('status 403')) {
      return {
        content: [{
          type: "text",
          text: "Draft creation was denied by Microsoft Graph (403). The token likely lacks Mail.ReadWrite scope. Re-authenticate with force=true to refresh consent, then try again."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error creating draft email: ${error.message}`
      }]
    };
  }
}

module.exports = handleDraftEmail;
