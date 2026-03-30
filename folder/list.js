/**
 * List folders functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List folders handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListFolders(args) {
  const includeItemCounts = args.includeItemCounts === true;
  const includeChildren = args.includeChildren === true;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Get all mail folders
    const folders = await getAllFoldersHierarchy(accessToken, includeItemCounts);
    
    // If including children, format as hierarchy
    if (includeChildren) {
      return {
        content: [{ 
          type: "text", 
          text: formatFolderHierarchy(folders, includeItemCounts)
        }]
      };
    } else {
      // Otherwise, format as flat list
      return {
        content: [{ 
          type: "text", 
          text: formatFolderList(folders, includeItemCounts)
        }]
      };
    }
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
        text: `Error listing folders: ${error.message}`
      }]
    };
  }
}

/**
 * Get all mail folders with hierarchy information
 * @param {string} accessToken - Access token
 * @param {boolean} includeItemCounts - Include item counts in response
 * @returns {Promise<Array>} - Array of folder objects with hierarchy
 */
async function getAllFoldersHierarchy(accessToken, includeItemCounts) {
  try {
    const selectFields = includeItemCounts
      ? 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      : 'id,displayName,parentFolderId,childFolderCount';

    const allFolders = await fetchFoldersRecursive(accessToken, 'me/mailFolders', selectFields);

    // Mark top-level folders (those fetched from the root endpoint)
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $top: 100, $select: 'id' }
    );
    const topLevelIds = new Set((response.value || []).map(f => f.id));

    return allFolders.map(folder => ({
      ...folder,
      isTopLevel: topLevelIds.has(folder.id)
    }));
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    throw error;
  }
}

/**
 * Recursively fetch folders and all their descendants
 */
async function fetchFoldersRecursive(accessToken, endpoint, selectFields) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    { $top: 100, $select: selectFields }
  );

  if (!response.value || response.value.length === 0) {
    return [];
  }

  const folders = response.value;
  const withChildren = folders.filter(f => f.childFolderCount > 0);

  const childResults = await Promise.all(withChildren.map(async (folder) => {
    try {
      return await fetchFoldersRecursive(
        accessToken,
        `me/mailFolders/${folder.id}/childFolders`,
        selectFields
      );
    } catch (error) {
      console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
      return [];
    }
  }));

  return [...folders, ...childResults.flat()];
}

/**
 * Format folders as a flat list
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted list
 */
function formatFolderList(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }
  
  // Sort folders alphabetically, with well-known folders first
  const wellKnownFolderNames = ['Inbox', 'Drafts', 'Sent Items', 'Deleted Items', 'Junk Email', 'Archive'];
  
  const sortedFolders = [...folders].sort((a, b) => {
    // Well-known folders come first
    const aIsWellKnown = wellKnownFolderNames.includes(a.displayName);
    const bIsWellKnown = wellKnownFolderNames.includes(b.displayName);
    
    if (aIsWellKnown && !bIsWellKnown) return -1;
    if (!aIsWellKnown && bIsWellKnown) return 1;
    
    if (aIsWellKnown && bIsWellKnown) {
      // Sort well-known folders by their index in the array
      return wellKnownFolderNames.indexOf(a.displayName) - wellKnownFolderNames.indexOf(b.displayName);
    }
    
    // Sort other folders alphabetically
    return a.displayName.localeCompare(b.displayName);
  });
  
  // Format each folder
  const folderLines = sortedFolders.map(folder => {
    let folderInfo = folder.displayName;
    
    // Add parent folder info if available
    if (folder.parentFolder) {
      folderInfo += ` (in ${folder.parentFolder})`;
    }
    
    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      folderInfo += ` - ${totalCount} items`;
      
      if (unreadCount > 0) {
        folderInfo += ` (${unreadCount} unread)`;
      }
    }
    
    return folderInfo;
  });
  
  return `Found ${folders.length} folders:\n\n${folderLines.join('\n')}`;
}

/**
 * Format folders as a hierarchical tree
 * @param {Array} folders - Array of folder objects
 * @param {boolean} includeItemCounts - Whether to include item counts
 * @returns {string} - Formatted hierarchy
 */
function formatFolderHierarchy(folders, includeItemCounts) {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }
  
  // Build folder hierarchy
  const folderMap = new Map();
  const rootFolders = [];
  
  // First pass: create map of all folders
  folders.forEach(folder => {
    folderMap.set(folder.id, {
      ...folder,
      children: []
    });
    
    if (folder.isTopLevel) {
      rootFolders.push(folder.id);
    }
  });
  
  // Second pass: build hierarchy
  folders.forEach(folder => {
    if (!folder.isTopLevel && folder.parentFolderId) {
      const parent = folderMap.get(folder.parentFolderId);
      if (parent) {
        parent.children.push(folder.id);
      } else {
        // Fallback for orphaned folders
        rootFolders.push(folder.id);
      }
    }
  });
  
  // Format hierarchy recursively
  function formatSubtree(folderId, level = 0) {
    const folder = folderMap.get(folderId);
    if (!folder) return '';
    
    const indent = '  '.repeat(level);
    let line = `${indent}${folder.displayName}`;
    
    // Add item counts if requested
    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      line += ` - ${totalCount} items`;
      
      if (unreadCount > 0) {
        line += ` (${unreadCount} unread)`;
      }
    }
    
    // Add children
    const childLines = folder.children
      .map(childId => formatSubtree(childId, level + 1))
      .filter(line => line.length > 0)
      .join('\n');
    
    return childLines.length > 0 ? `${line}\n${childLines}` : line;
  }
  
  // Format all root folders
  const formattedHierarchy = rootFolders
    .map(folderId => formatSubtree(folderId))
    .join('\n');
  
  return `Folder Hierarchy:\n\n${formattedHierarchy}`;
}

module.exports = handleListFolders;
