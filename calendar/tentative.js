/**
 * Tentatively accept event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Tentatively accept event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleTentativeEvent(args) {
  const { eventId, comment } = args;

  if (!eventId) {
    return {
      content: [{
        type: "text",
        text: "Event ID is required to tentatively accept an event."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build API endpoint
    const endpoint = `me/events/${eventId}/tentativelyAccept`;

    // Request body
    const body = {
      comment: comment || "Tentatively accepted via API"
    };

    // Make API call
    await callGraphAPI(accessToken, 'POST', endpoint, body);

    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} has been tentatively accepted.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
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
        text: `Error tentatively accepting event: ${error.message}`
      }]
    };
  }
}

module.exports = handleTentativeEvent;
