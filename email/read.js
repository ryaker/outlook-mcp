/**
 * Read email functionality
 *
 * Security: HTML emails are sanitized to remove hidden content that could
 * be used for prompt injection attacks. Only visible text is extracted.
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { processHtmlEmail, sanitizeHtmlToText } = require('../utils/html-sanitizer');

/**
 * Read email handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @param {boolean} args.includeRawHtml - If true, include raw HTML (unsafe, for debugging only)
 * @returns {object} - MCP response
 */
async function handleReadEmail(args) {
  const emailId = args.id;
  const includeRawHtml = args.includeRawHtml === true;

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Make API call to get email details
    const endpoint = `me/messages/${encodeURIComponent(emailId)}`;
    const queryParams = {
      $select: config.EMAIL_DETAIL_FIELDS
    };

    try {
      const email = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

      if (!email) {
        return {
          content: [
            {
              type: "text",
              text: `Email with ID ${emailId} not found.`
            }
          ]
        };
      }

      // Format sender, recipients, etc.
      const sender = email.from ? `${email.from.emailAddress.name} (${email.from.emailAddress.address})` : 'Unknown';
      const senderAddress = email.from?.emailAddress?.address || 'unknown';
      const to = email.toRecipients ? email.toRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const cc = email.ccRecipients && email.ccRecipients.length > 0 ? email.ccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const bcc = email.bccRecipients && email.bccRecipients.length > 0 ? email.bccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const date = new Date(email.receivedDateTime).toLocaleString();

      // Extract and sanitize body content
      let body = '';
      let bodyNote = '';

      if (email.body) {
        if (email.body.contentType === 'html') {
          // Use secure HTML sanitizer to extract visible text only
          // This prevents prompt injection via hidden HTML content
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: senderAddress,
              subject: email.subject,
              date: date
            }
          });
          bodyNote = '\n[HTML email - sanitized for security, hidden content removed]\n';
        } else {
          // Plain text - still wrap with boundary for safety
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: senderAddress,
              subject: email.subject,
              date: date
            }
          });
        }
      } else {
        body = email.bodyPreview || 'No content';
      }

      // Format the email
      const formattedEmail = `From: ${sender}
To: ${to}
${cc !== 'None' ? `CC: ${cc}\n` : ''}${bcc !== 'None' ? `BCC: ${bcc}\n` : ''}Subject: ${email.subject}
Date: ${date}
Importance: ${email.importance || 'normal'}
Has Attachments: ${email.hasAttachments ? 'Yes' : 'No'}
${bodyNote}
${body}`;

      // Optionally include raw HTML for debugging (not recommended for normal use)
      let rawHtmlSection = '';
      if (includeRawHtml && email.body?.contentType === 'html') {
        rawHtmlSection = `\n\n--- RAW HTML (UNSAFE - FOR DEBUGGING ONLY) ---\n${email.body.content}\n--- END RAW HTML ---`;
      }

      return {
        content: [
          {
            type: "text",
            text: formattedEmail + rawHtmlSection
          }
        ]
      };
    } catch (error) {
      console.error(`Error reading email: ${error.message}`);
      
      // Improved error handling with more specific messages
      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [
            {
              type: "text",
              text: `The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.`
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: `Failed to read email: ${error.message}`
            }
          ]
        };
      }
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error accessing email: ${error.message}`
      }]
    };
  }
}

module.exports = handleReadEmail;
