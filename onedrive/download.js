/**
 * OneDrive get download URL functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Get download URL handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDownload(args) {
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

    // Build endpoint - by ID or by path
    let endpoint;
    if (itemId) {
      endpoint = `me/drive/items/${itemId}`;
    } else {
      const normalizedPath = path.replace(/^\/+|\/+$/g, '');
      endpoint = `me/drive/root:/${normalizedPath}`;
    }

    // Get item metadata with download URL
    const queryParams = {
      $select: 'id,name,size,@microsoft.graph.downloadUrl'
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response) {
      return {
        content: [{
          type: "text",
          text: "File not found."
        }]
      };
    }

    const downloadUrl = response['@microsoft.graph.downloadUrl'];

    if (!downloadUrl) {
      // If no direct download URL, this might be a folder
      if (response.folder) {
        return {
          content: [{
            type: "text",
            text: `"${response.name}" is a folder and cannot be downloaded directly.`
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: "Could not get download URL for this item."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Download URL for "${response.name}" (${formatSize(response.size)}):\n\n${downloadUrl}\n\nNote: This URL is pre-authenticated and expires after a short time.`
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
        text: `Error getting download URL: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleDownload;
