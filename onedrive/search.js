/**
 * OneDrive search files functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Search files handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchFiles(args) {
  const query = args.query;
  const count = args.count || 25;

  if (!query) {
    return {
      content: [{
        type: "text",
        text: "Search query is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Use the search endpoint
    const endpoint = `me/drive/search(q='${encodeURIComponent(query)}')`;

    const queryParams = {
      $top: Math.min(50, count),
      $select: config.ONEDRIVE_SELECT_FIELDS
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No files found matching "${query}".`
        }]
      };
    }

    // Format results
    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = new Date(item.lastModifiedDateTime).toLocaleString();
      const path = item.parentReference?.path?.replace('/drive/root:', '') || '/';

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Path: ${path}\n   Modified: ${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} items matching "${query}":\n\n${fileList}`
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
        text: `Error searching files: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleSearchFiles;
