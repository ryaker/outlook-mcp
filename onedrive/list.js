/**
 * OneDrive list files/folders functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List files handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFiles(args) {
  const path = args.path || '';
  const count = args.count || 25;

  try {
    const accessToken = await ensureAuthenticated();

    // Build endpoint - root or specific path
    let endpoint;
    if (!path || path === '/' || path === 'root') {
      endpoint = 'me/drive/root/children';
    } else {
      // Normalize path - remove leading/trailing slashes
      const normalizedPath = path.replace(/^\/+|\/+$/g, '');
      endpoint = `me/drive/root:/${normalizedPath}:/children`;
    }

    const queryParams = {
      $top: Math.min(50, count),
      $select: config.ONEDRIVE_SELECT_FIELDS,
      $orderby: 'name'
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No files found in ${path || 'root'}.`
        }]
      };
    }

    // Format results
    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = new Date(item.lastModifiedDateTime).toLocaleString();

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Modified: ${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} items in ${path || 'root'}:\n\n${fileList}`
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
        text: `Error listing files: ${error.message}`
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

module.exports = handleListFiles;
