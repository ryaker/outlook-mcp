/**
 * OneDrive folder operations (create/delete)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create folder handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const path = args.path;
  const name = args.name;

  if (!name) {
    return {
      content: [{
        type: "text",
        text: "Folder name is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Build parent folder endpoint
    let endpoint;
    if (!path || path === '/' || path === 'root') {
      endpoint = 'me/drive/root/children';
    } else {
      const normalizedPath = path.replace(/^\/+|\/+$/g, '');
      endpoint = `me/drive/root:/${normalizedPath}:/children`;
    }

    const body = {
      name: name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    };

    const response = await callGraphAPI(accessToken, 'POST', endpoint, body);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to create folder."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully created folder "${response.name}"\n\nID: ${response.id}\nWeb URL: ${response.webUrl}`
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
        text: `Error creating folder: ${error.message}`
      }]
    };
  }
}

/**
 * Delete item handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteItem(args) {
  const itemId = args.itemId;
  const path = args.path;

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

    // Get item details first (to confirm existence and get name)
    let endpoint;
    if (itemId) {
      endpoint = `me/drive/items/${itemId}`;
    } else {
      const normalizedPath = path.replace(/^\/+|\/+$/g, '');
      endpoint = `me/drive/root:/${normalizedPath}`;
    }

    // Get item info first
    const itemInfo = await callGraphAPI(accessToken, 'GET', endpoint);

    if (!itemInfo || !itemInfo.id) {
      return {
        content: [{
          type: "text",
          text: "Item not found."
        }]
      };
    }

    const itemName = itemInfo.name;
    const isFolder = !!itemInfo.folder;

    // Delete the item
    const deleteEndpoint = `me/drive/items/${itemInfo.id}`;
    await callGraphAPI(accessToken, 'DELETE', deleteEndpoint);

    return {
      content: [{
        type: "text",
        text: `Successfully deleted ${isFolder ? 'folder' : 'file'} "${itemName}".`
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
        text: `Error deleting item: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleCreateFolder,
  handleDeleteItem
};
