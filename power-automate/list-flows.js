/**
 * Power Automate list flows functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');

/**
 * List flows handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFlows(args) {
  const environmentId = args.environmentId;

  if (!environmentId) {
    return {
      content: [{
        type: "text",
        text: "Environment ID is required. Use 'flow-list-environments' to get available environments."
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

    const path = `/environments/${environmentId}/flows`;
    const response = await callFlowAPI(accessToken, 'GET', path);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No flows found in environment ${environmentId}.\n\nNote: Only solution-aware flows are accessible via the API.`
        }]
      };
    }

    const flowList = response.value.map((flow, index) => {
      const props = flow.properties || {};
      const state = props.state || 'Unknown';
      const stateIcon = state === 'Started' ? '[ON]' : '[OFF]';
      const triggerType = props.definition?.triggers ? Object.keys(props.definition.triggers)[0] : 'Unknown';
      const created = props.createdTime ? new Date(props.createdTime).toLocaleDateString() : 'Unknown';

      return `${index + 1}. ${stateIcon} ${props.displayName || flow.name}\n   ID: ${flow.name}\n   Trigger: ${triggerType}\n   Created: ${created}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} flow(s) in environment:\n\n${flowList}`
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

    return {
      content: [{
        type: "text",
        text: `Error listing flows: ${error.message}`
      }]
    };
  }
}

module.exports = handleListFlows;
