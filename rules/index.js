/**
 * Email rules management module for Outlook MCP server
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const handleListRules = require('./list');
const handleCreateRule = require('./create');
const handleDeleteRule = require('./delete');

// Import getInboxRules for the edit sequence tool
const { getInboxRules } = require('./list');

/**
 * Edit rule sequence handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;
  
  if (!ruleName) {
    return {
      content: [{ 
        type: "text", 
        text: "Rule name is required. Please specify the exact name of an existing rule."
      }]
    };
  }
  
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return {
      content: [{ 
        type: "text", 
        text: "A positive sequence number is required. Lower numbers run first (higher priority)."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get all rules
    const rules = await getInboxRules(accessToken);
    
    // Find the rule by name
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return {
        content: [{ 
          type: "text", 
          text: `Rule with name "${ruleName}" not found.`
        }]
      };
    }
    
    // Update the rule sequence
    const updateResult = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/mailFolders/inbox/messageRules/${rule.id}`,
      {
        sequence: sequence
      }
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Successfully updated the sequence of rule "${ruleName}" to ${sequence}.`
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
        text: `Error updating rule sequence: ${error.message}`
      }]
    };
  }
}

// Rules management tool definitions
const rulesTools = [
  {
    name: "list-rules",
    description: "Lists inbox rules in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {
        includeDetails: {
          type: "boolean",
          description: "Include detailed rule conditions and actions"
        }
      },
      required: []
    },
    handler: handleListRules
  },
  {
    name: "create-rule",
    description: "Creates a new inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        name: {
          type: "string",
          description: "Name of the rule to create"
        },
        fromAddresses: {
          type: "string",
          description: "Comma-separated list of sender email addresses for the rule"
        },
        containsSubject: {
          type: "string",
          description: "Subject text the email must contain"
        },
        hasAttachments: {
          type: "boolean",
          description: "Whether the rule applies to emails with attachments"
        },
        moveToFolder: {
          type: "string",
          description: "Name of the folder to move matching emails to"
        },
        markAsRead: {
          type: "boolean", 
          description: "Whether to mark matching emails as read"
        },
        isEnabled: {
          type: "boolean",
          description: "Whether the rule should be enabled after creation (default: true)"
        },
        sequence: {
          type: "number",
          description: "Order in which the rule is executed (lower numbers run first, default: 100)"
        }
      },
      required: ["name"]
    },
    handler: handleCreateRule
  },
  {
    name: "edit-rule-sequence",
    description: "Changes the execution order of an existing inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: {
          type: "string",
          description: "Name of the rule to modify"
        },
        sequence: {
          type: "number",
          description: "New sequence value for the rule (lower numbers run first)"
        }
      },
      required: ["ruleName", "sequence"]
    },
    handler: handleEditRuleSequence
  },
  {
    name: "delete-rule",
    description: "Deletes an existing inbox rule by name or ID. If multiple rules share the same name, the tool returns a list and asks you to re-call with a specific ruleId.",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: {
          type: "string",
          description: "Name of the rule to delete. If multiple rules share this name, the tool will return the list and ask for a specific ID."
        },
        ruleId: {
          type: "string",
          description: "Specific rule ID (from Microsoft Graph) to delete. Use this when multiple rules share the same name."
        }
      },
      required: []
    },
    handler: handleDeleteRule
  }
];

module.exports = {
  rulesTools,
  handleListRules,
  handleCreateRule,
  handleDeleteRule,
  handleEditRuleSequence
};
