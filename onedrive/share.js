/**
 * OneDrive create sharing link functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create sharing link handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleShare(args) {
  const itemId = args.itemId;
  const path = args.path;
  const type = args.type || 'view'; // view, edit, embed
  const scope = args.scope || 'anonymous'; // anonymous, organization

  if (!itemId && !path) {
    return {
      content: [{
        type: "text",
        text: "Either itemId or path is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // First get the item ID if we only have path
    let resolvedItemId = itemId;
    let itemName = '';

    if (!resolvedItemId && path) {
      const normalizedPath = path.replace(/^\/+|\/+$/g, '');
      const itemEndpoint = `me/drive/root:/${normalizedPath}`;
      const itemResponse = await callGraphAPI(accessToken, 'GET', itemEndpoint);

      if (!itemResponse || !itemResponse.id) {
        return {
          content: [{
            type: "text",
            text: `File not found at path: ${path}`
          }]
        };
      }

      resolvedItemId = itemResponse.id;
      itemName = itemResponse.name;
    }

    // Create the sharing link
    const endpoint = `me/drive/items/${resolvedItemId}/createLink`;
    const body = {
      type: type,
      scope: scope
    };

    const response = await callGraphAPI(accessToken, 'POST', endpoint, body);

    if (!response || !response.link) {
      return {
        content: [{
          type: "text",
          text: "Failed to create sharing link."
        }]
      };
    }

    const linkInfo = response.link;
    const shareText = itemName
      ? `Sharing link created for "${itemName}":`
      : `Sharing link created:`;

    return {
      content: [{
        type: "text",
        text: `${shareText}\n\nLink: ${linkInfo.webUrl}\nType: ${type}\nScope: ${scope}\n\nNote: ${scope === 'anonymous' ? 'Anyone with this link can access the file.' : 'Only people in your organization can access.'}`
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
        text: `Error creating sharing link: ${error.message}`
      }]
    };
  }
}

module.exports = handleShare;
