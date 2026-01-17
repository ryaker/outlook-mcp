/**
 * Power Automate run/trigger flow functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');

/**
 * Run flow handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRunFlow(args) {
  const environmentId = args.environmentId;
  const flowId = args.flowId;
  const inputs = args.inputs;

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

    // Trigger the flow
    const path = `/environments/${environmentId}/flows/${flowId}/triggers/manual/run`;

    // Parse inputs if provided as JSON string
    let inputData = null;
    if (inputs) {
      try {
        inputData = typeof inputs === 'string' ? JSON.parse(inputs) : inputs;
      } catch (e) {
        return {
          content: [{
            type: "text",
            text: `Invalid inputs format. Please provide valid JSON: ${e.message}`
          }]
        };
      }
    }

    const response = await callFlowAPI(accessToken, 'POST', path, inputData);

    // The response might include run details or just acknowledgment
    const runId = response.name || response.id || 'initiated';

    return {
      content: [{
        type: "text",
        text: `Flow triggered successfully!\n\nRun ID: ${runId}\n\nUse 'flow-list-runs' to check the status of this run.`
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
          text: "Cannot trigger this flow. Ensure:\n1. The flow has a 'manual' trigger\n2. The flow is enabled (state: Started)\n3. You have permission to run this flow"
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error running flow: ${error.message}`
      }]
    };
  }
}

module.exports = handleRunFlow;
