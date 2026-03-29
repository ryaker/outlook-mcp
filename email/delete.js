/**
 * Delete email functionality (move to Deleted Items / trash)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Delete email handler - moves email to Deleted Items folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteEmail(args) {
  const emailId = args.id;
  const permanent = args.permanent === true;

  if (!emailId) {
    return {
      content: [{ type: "text", text: "Email ID is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    if (permanent) {
      // Hard delete - permanently remove the message
      await callGraphAPI(accessToken, 'DELETE', `me/messages/${encodeURIComponent(emailId)}`);
      return {
        content: [{ type: "text", text: "Email permanently deleted." }]
      };
    } else {
      // Soft delete - move to Deleted Items (trash)
      const result = await callGraphAPI(accessToken, 'POST', `me/messages/${encodeURIComponent(emailId)}/move`, {
        destinationId: 'deleteditems'
      });
      return {
        content: [{ type: "text", text: `Email moved to Deleted Items. ID: ${result.id}` }]
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Failed to delete email: ${error.message}` }]
    };
  }
}

module.exports = handleDeleteEmail;
