/**
 * Power Automate toggle flow (enable/disable) functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');

/**
 * Toggle flow handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleToggleFlow(args) {
  const environmentId = args.environmentId;
  const flowId = args.flowId;
  const enable = args.enable !== false; // Default to enable if not specified

  if (!environmentId || !flowId) {
    return {
      content: [{
        type: "text",
        text: "Both environmentId and flowId are required."
      }]
    };
  }

  try {
    const accessToken = getFlowAccessToken();

    if (!accessToken) {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication required. Please authenticate with Flow scope first."
        }]
      };
    }

    // The action endpoint depends on whether we're enabling or disabling
    const action = enable ? 'start' : 'stop';
    const path = `/environments/${environmentId}/flows/${flowId}/${action}`;

    await callFlowAPI(accessToken, 'POST', path);

    const actionText = enable ? 'enabled' : 'disabled';

    return {
      content: [{
        type: "text",
        text: `Flow successfully ${actionText}.\n\nFlow ID: ${flowId}\nNew State: ${enable ? 'Started' : 'Stopped'}`
      }]
    };
  } catch (error) {
    if (error.message === 'FLOW_UNAUTHORIZED') {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication expired. Please re-authenticate with Flow scope."
        }]
      };
    }

    if (error.message.includes('403')) {
      return {
        content: [{
          type: "text",
          text: "Cannot modify this flow. Ensure you have owner or editor permissions."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error toggling flow: ${error.message}`
      }]
    };
  }
}

module.exports = handleToggleFlow;
