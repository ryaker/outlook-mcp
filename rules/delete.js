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
  const { ruleName, ruleId } = args;

  if (!ruleName && !ruleId) {
    return {
      content: [{
        type: "text",
        text: "Either ruleName or ruleId is required. Provide the exact name of an existing rule or its Graph API ID."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    let targetRuleId = ruleId;
    let targetRuleName = ruleName;

    // If only ruleName was provided, resolve it to an ID
    if (!targetRuleId) {
      const rules = await getInboxRules(accessToken);
      const matchingRules = rules.filter(r => r.displayName === ruleName);

      if (matchingRules.length === 0) {
        return {
          content: [{
            type: "text",
            text: `No rule found with name "${ruleName}". Use 'list-rules' to see all rules.`
          }]
        };
      }

      if (matchingRules.length > 1) {
        const ruleList = matchingRules
          .map((r, i) => {
            const conditions = [];
            if (r.conditions?.fromAddresses?.length > 0) {
              conditions.push(`from: ${r.conditions.fromAddresses.map(a => a.emailAddress.address).join(', ')}`);
            }
            if (r.conditions?.subjectContains?.length > 0) {
              conditions.push(`subject contains: "${r.conditions.subjectContains.join(', ')}"`);
            }
            if (r.conditions?.hasAttachment === true) {
              conditions.push('has attachment');
            }
            const conditionsText = conditions.length > 0 ? ` [${conditions.join('; ')}]` : ' [no conditions]';
            return `${i + 1}. ID: ${r.id}${conditionsText}`;
          })
          .join('\n');

        return {
          content: [{
            type: "text",
            text: `Multiple rules found with name "${ruleName}". Call 'delete-rule' again with a specific ruleId:\n\n${ruleList}`
          }]
        };
      }

      targetRuleId = matchingRules[0].id;
      targetRuleName = matchingRules[0].displayName;
    }

    // Delete the rule via Microsoft Graph
    await callGraphAPI(
      accessToken,
      'DELETE',
      `me/mailFolders/inbox/messageRules/${targetRuleId}`,
      null
    );

    return {
      content: [{
        type: "text",
        text: `Successfully deleted rule${targetRuleName ? ` "${targetRuleName}"` : ''} (ID: ${targetRuleId}).`
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
