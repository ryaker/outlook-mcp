/**
 * Create rule functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { getFolderIdByName } = require('../email/folder-utils');
const { getInboxRules } = require('./list');

/**
 * Create rule handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateRule(args) {
  const {
    name,
    fromAddresses,
    containsSubject,
    bodyContains,
    bodyOrSubjectContains,
    sentToAddresses,
    hasAttachments,
    moveToFolder,
    markAsRead,
    isEnabled = true,
    sequence
  } = args;
  
  // Add validation for sequence parameter
  if (sequence !== undefined && (isNaN(sequence) || sequence < 1)) {
    return {
      content: [{ 
        type: "text", 
        text: "Sequence must be a positive number greater than zero."
      }]
    };
  }
  
  if (!name) {
    return {
      content: [{ 
        type: "text", 
        text: "Rule name is required."
      }]
    };
  }
  
  // Validate that at least one condition or action is specified
  const hasCondition = fromAddresses || containsSubject || bodyContains || bodyOrSubjectContains || sentToAddresses || hasAttachments === true;
  const hasAction = moveToFolder || markAsRead === true;
  
  if (!hasCondition) {
    return {
      content: [{ 
        type: "text", 
        text: "At least one condition is required. Specify fromAddresses, containsSubject, or hasAttachments."
      }]
    };
  }
  
  if (!hasAction) {
    return {
      content: [{ 
        type: "text", 
        text: "At least one action is required. Specify moveToFolder or markAsRead."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Create rule
    const result = await createInboxRule(accessToken, {
      name,
      fromAddresses,
      containsSubject,
      bodyContains,
      bodyOrSubjectContains,
      sentToAddresses,
      hasAttachments,
      moveToFolder,
      markAsRead,
      isEnabled,
      sequence
    });

    let responseText = result.message;
    
    // Add a tip about sequence if it wasn't provided
    if (!sequence && !result.error) {
      responseText += "\n\nTip: You can specify a 'sequence' parameter when creating rules to control their execution order. Lower sequence numbers run first.";
    }
    
    return {
      content: [{ 
        type: "text", 
        text: responseText
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
        text: `Error creating rule: ${error.message}`
      }]
    };
  }
}

/**
 * Create a new inbox rule
 * @param {string} accessToken - Access token
 * @param {object} ruleOptions - Rule creation options
 * @returns {Promise<object>} - Result object with status and message
 */
async function createInboxRule(accessToken, ruleOptions) {
  try {
    const {
      name,
      fromAddresses,
      containsSubject,
      bodyContains,
      bodyOrSubjectContains,
      sentToAddresses,
      hasAttachments,
      moveToFolder,
      markAsRead,
      isEnabled,
      sequence
    } = ruleOptions;
    
    // Get existing rules to determine sequence if not provided
    let ruleSequence = sequence;
    if (!ruleSequence) {
      try {
        // Default to 100 if we can't get existing rules
        ruleSequence = 100;
        
        // Get existing rules to find highest sequence
        const existingRules = await getInboxRules(accessToken);
        if (existingRules && existingRules.length > 0) {
          // Find the highest sequence
          const highestSequence = Math.max(...existingRules.map(r => r.sequence || 0));
          // Set new rule sequence to be higher
          ruleSequence = Math.max(highestSequence + 1, 100);
          console.error(`Auto-generated sequence: ${ruleSequence} (based on highest existing: ${highestSequence})`);
        }
      } catch (sequenceError) {
        console.error(`Error determining rule sequence: ${sequenceError.message}`);
        // Fall back to default value
        ruleSequence = 100;
      }
    }
    
    console.error(`Using rule sequence: ${ruleSequence}`);
    
    // Make sure sequence is a positive integer
    ruleSequence = Math.max(1, Math.floor(ruleSequence));
    
    // Build rule object with sequence
    const rule = {
      displayName: name,
      isEnabled: isEnabled === true,
      sequence: ruleSequence,
      conditions: {},
      actions: {}
    };
    
    // Add conditions
    if (fromAddresses) {
      // Parse email addresses
      const emailAddresses = fromAddresses.split(',')
        .map(email => email.trim())
        .filter(email => email)
        .map(email => ({
          emailAddress: {
            address: email
          }
        }));
      
      if (emailAddresses.length > 0) {
        rule.conditions.fromAddresses = emailAddresses;
      }
    }
    
    if (containsSubject) {
      rule.conditions.subjectContains = [containsSubject];
    }

    if (bodyContains) {
      rule.conditions.bodyContains = [bodyContains];
    }

    if (bodyOrSubjectContains) {
      rule.conditions.bodyOrSubjectContains = [bodyOrSubjectContains];
    }

    if (sentToAddresses) {
      const toAddresses = sentToAddresses.split(',')
        .map(email => email.trim())
        .filter(email => email)
        .map(email => ({ emailAddress: { address: email } }));
      if (toAddresses.length > 0) {
        rule.conditions.sentToAddresses = toAddresses;
      }
    }

    if (hasAttachments === true) {
      rule.conditions.hasAttachment = true;
    }
    
    // Add actions
    if (moveToFolder) {
      // Get folder ID
      try {
        const folderId = await getFolderIdByName(accessToken, moveToFolder);
        if (!folderId) {
          return {
            success: false,
            message: `Target folder "${moveToFolder}" not found. Please specify a valid folder name.`
          };
        }
        
        rule.actions.moveToFolder = folderId;
      } catch (folderError) {
        console.error(`Error resolving folder "${moveToFolder}": ${folderError.message}`);
        return {
          success: false,
          message: `Error resolving folder "${moveToFolder}": ${folderError.message}`
        };
      }
    }
    
    if (markAsRead === true) {
      rule.actions.markAsRead = true;
    }
    
    // Create the rule
    const response = await callGraphAPI(
      accessToken,
      'POST',
      'me/mailFolders/inbox/messageRules',
      rule
    );
    
    if (response && response.id) {
      return {
        success: true,
        message: `Successfully created rule "${name}" with sequence ${ruleSequence}.`,
        ruleId: response.id
      };
    } else {
      return {
        success: false,
        message: "Failed to create rule. The server didn't return a rule ID."
      };
    }
  } catch (error) {
    console.error(`Error creating rule: ${error.message}`);
    throw error;
  }
}

module.exports = handleCreateRule;
