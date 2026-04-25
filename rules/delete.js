/**
 * Delete rule functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getInboxRules } = require('./list');

/**
 * Delete rule handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteRule(args) {
  const { ruleName } = args;

  if (!ruleName) {
    return {
      content: [{
        type: "text",
        text: "Rule name is required. Use 'list-rules' to see existing rules."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Find the rule by name
    const rules = await getInboxRules(accessToken);
    const rule = rules.find(r => r.displayName === ruleName);

    if (!rule) {
      return {
        content: [{
          type: "text",
          text: `Rule "${ruleName}" not found. Use 'list-rules' to see existing rules.`
        }]
      };
    }

    // Delete the rule
    await callGraphAPI(
      accessToken,
      'DELETE',
      `me/mailFolders/inbox/messageRules/${rule.id}`,
      null
    );

    return {
      content: [{
        type: "text",
        text: `Successfully deleted rule "${ruleName}".`
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

    return {
      content: [{
        type: "text",
        text: `Error deleting rule: ${error.message}`
      }]
    };
  }
}

module.exports = handleDeleteRule;
