/**
 * Power Automate list environments functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');

/**
 * List environments handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEnvironments(args) {
  try {
    const accessToken = getFlowAccessToken();

    if (!accessToken) {
      return {
        content: [{
          type: "text",
          text: "Power Automate authentication required. Please authenticate with Flow scope first.\n\nNote: Power Automate requires additional Azure AD configuration with the Flow API scope."
        }]
      };
    }

    const response = await callFlowAPI(accessToken, 'GET', '/environments');

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No Power Platform environments found."
        }]
      };
    }

    const envList = response.value.map((env, index) => {
      const props = env.properties || {};
      const isDefault = props.isDefault ? ' [DEFAULT]' : '';
      const region = props.azureRegionHint || 'Unknown region';

      return `${index + 1}. ${props.displayName || env.name}${isDefault}\n   ID: ${env.name}\n   Region: ${region}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} environment(s):\n\n${envList}\n\nUse the environment ID (e.g., 'Default-12345') with other flow commands.`
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
        text: `Error listing environments: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEnvironments;
