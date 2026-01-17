/**
 * Power Automate module for M365 MCP server
 */
const handleListEnvironments = require('./list-environments');
const handleListFlows = require('./list-flows');
const handleRunFlow = require('./run-flow');
const handleListRuns = require('./list-runs');
const handleToggleFlow = require('./toggle-flow');

// Power Automate tool definitions
const powerAutomateTools = [
  {
    name: "flow-list-environments",
    description: "List available Power Platform environments",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListEnvironments
  },
  {
    name: "flow-list",
    description: "List flows in a Power Platform environment",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "The environment ID (e.g., 'Default-12345'). Use 'flow-list-environments' to find available environments."
        }
      },
      required: ["environmentId"]
    },
    handler: handleListFlows
  },
  {
    name: "flow-run",
    description: "Trigger a manual flow run",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "The environment ID"
        },
        flowId: {
          type: "string",
          description: "The flow ID to trigger"
        },
        inputs: {
          type: "string",
          description: "Optional JSON string of input parameters for the flow"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleRunFlow
  },
  {
    name: "flow-list-runs",
    description: "Get run history for a flow",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "The environment ID"
        },
        flowId: {
          type: "string",
          description: "The flow ID to get runs for"
        },
        count: {
          type: "number",
          description: "Number of runs to retrieve (default: 10)"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleListRuns
  },
  {
    name: "flow-toggle",
    description: "Enable or disable a flow",
    inputSchema: {
      type: "object",
      properties: {
        environmentId: {
          type: "string",
          description: "The environment ID"
        },
        flowId: {
          type: "string",
          description: "The flow ID to toggle"
        },
        enable: {
          type: "boolean",
          description: "Set to true to enable, false to disable (default: true)"
        }
      },
      required: ["environmentId", "flowId"]
    },
    handler: handleToggleFlow
  }
];

module.exports = {
  powerAutomateTools,
  handleListEnvironments,
  handleListFlows,
  handleRunFlow,
  handleListRuns,
  handleToggleFlow
};
