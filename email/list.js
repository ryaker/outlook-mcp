/**
 * List emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";

  // Parse and validate count with proper error handling
  const parsedCount = parseInt(args.count, 10);
  const defaultCount = 10;
  const validCount = (Number.isNaN(parsedCount) || parsedCount <= 0) ? defaultCount : parsedCount;

  // Cap to reasonable maximum to prevent excessive API calls
  const requestedCount = Math.min(validCount, config.MAX_TOTAL_RESULTS);

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build API endpoint
    let endpoint = 'me/messages';
    if (folder.toLowerCase() !== 'inbox') {
      // Get folder ID first if not inbox
      const folderResponse = await callGraphAPI(
        accessToken,
        'GET',
        `me/mailFolders?$filter=displayName eq '${folder}'`
      );

      if (folderResponse.value && folderResponse.value.length > 0) {
        endpoint = `me/mailFolders/${folderResponse.value[0].id}/messages`;
      }
    }

    // Add query parameters with safe page size (at least 1, capped at API_PAGE_SIZE)
    const queryParams = {
      $top: Math.max(1, Math.min(config.API_PAGE_SIZE, requestedCount)),
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };

    // Make API call with pagination support
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: `No emails found in ${folder}.`
        }]
      };
    }
    
    // Format results
    const emailList = response.value.map((email, index) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? '' : '[UNREAD] ';
      
      return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
    }).join("\n");
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} emails in ${folder}:\n\n${emailList}`
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
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;
