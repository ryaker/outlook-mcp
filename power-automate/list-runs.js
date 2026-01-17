/**
 * Power Automate list flow runs functionality
 */
const { callFlowAPI } = require('./flow-api');
const { getFlowAccessToken } = require('../auth/token-manager');

/**
 * List flow runs handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListRuns(args) {
  const environmentId = args.environmentId;
  const flowId = args.flowId;
  const count = args.count || 10;

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

    const path = `/environments/${environmentId}/flows/${flowId}/runs`;
    const response = await callFlowAPI(accessToken, 'GET', path);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No run history found for this flow."
        }]
      };
    }

    // Limit to requested count
    const runs = response.value.slice(0, count);

    const runList = runs.map((run, index) => {
      const props = run.properties || {};
      const status = props.status || 'Unknown';
      const statusIcon = getStatusIcon(status);
      const startTime = props.startTime ? new Date(props.startTime).toLocaleString() : 'Unknown';
      const endTime = props.endTime ? new Date(props.endTime).toLocaleString() : 'Running...';
      const duration = props.startTime && props.endTime
        ? formatDuration(new Date(props.endTime) - new Date(props.startTime))
        : 'N/A';

      return `${index + 1}. ${statusIcon} ${status}\n   Run ID: ${run.name}\n   Started: ${startTime}\n   Duration: ${duration}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Recent ${runs.length} run(s) for this flow:\n\n${runList}`
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
        text: `Error listing flow runs: ${error.message}`
      }]
    };
  }
}

/**
 * Get status icon for run status
 */
function getStatusIcon(status) {
  const icons = {
    'Succeeded': '[OK]',
    'Failed': '[FAIL]',
    'Running': '[...]',
    'Cancelled': '[X]',
    'TimedOut': '[TIMEOUT]'
  };
  return icons[status] || '[?]';
}

/**
 * Format duration in human-readable form
 */
function formatDuration(ms) {
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
  if (ms < 3600000) return `${Math.floor(ms / 60000)}m ${Math.floor((ms % 60000) / 1000)}s`;
  return `${Math.floor(ms / 3600000)}h ${Math.floor((ms % 3600000) / 60000)}m`;
}

module.exports = handleListRuns;
