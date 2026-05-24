/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {

  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  try {
    // Try to find the folder by name
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (folderId) {
      const path = `me/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, fall back to inbox
    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return WELL_KNOWN_FOLDERS['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return WELL_KNOWN_FOLDERS['inbox'];
  }
}

/**
 * Get the ID of a mail folder by its name
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  try {
    console.error(`Looking for folder with name "${folderName}"`);

    // Detect path-style input (e.g. "Archive/2024", "Inbox/To Delete") and
    // walk it segment-by-segment instead of treating the whole string as one name.
    const segments = folderName.split('/').map(s => s.trim()).filter(Boolean);
    if (segments.length > 1) {
      return await getFolderIdByPath(accessToken, segments);
    }

    // Single-segment: exact match filter first
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName.replace(/'/g, "''")}'` }
    );

    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }

    // Case-insensitive fallback across top-level folders
    console.error(`No exact match found for "${folderName}", trying case-insensitive search`);
    const allFoldersResponse = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $top: 100, $select: 'id,displayName,childFolderCount' }
    );

    if (allFoldersResponse.value) {
      const lowerFolderName = folderName.toLowerCase();
      const matchingFolder = allFoldersResponse.value.find(
        folder => folder.displayName.toLowerCase() === lowerFolderName
      );

      if (matchingFolder) {
        console.error(`Found case-insensitive match for "${folderName}" with ID: ${matchingFolder.id}`);
        return matchingFolder.id;
      }

      // Search one level of child folders in parallel
      const foldersWithChildren = allFoldersResponse.value.filter(f => f.childFolderCount > 0);
      const childResults = await Promise.all(foldersWithChildren.map(async (parent) => {
        try {
          const res = await callGraphAPI(
            accessToken,
            'GET',
            `me/mailFolders/${parent.id}/childFolders`,
            null,
            { $top: 100, $select: 'id,displayName,childFolderCount' }
          );
          return { parent, children: res.value || [] };
        } catch (err) {
          console.error(`Error searching child folders of "${parent.displayName}": ${err.message}`);
          return { parent, children: [] };
        }
      }));

      for (const { parent, children } of childResults) {
        const match = children.find(f => f.displayName.toLowerCase() === lowerFolderName);
        if (match) {
          console.error(`Found child folder "${folderName}" under "${parent.displayName}" with ID: ${match.id}`);
          return match.id;
        }
      }
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Walk a slash-separated folder path segment by segment.
 * @param {string} accessToken
 * @param {string[]} segments - path segments, e.g. ['Archive', '2024']
 * @returns {Promise<string|null>} - Folder ID of the final segment or null
 */
async function getFolderIdByPath(accessToken, segments) {
  console.error(`Resolving folder path: ${segments.join('/')}`);

  // Find the root segment among top-level folders
  const rootResponse = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders',
    null,
    { $top: 100, $select: 'id,displayName,childFolderCount' }
  );

  if (!rootResponse.value) {
    return null;
  }

  const rootName = segments[0].toLowerCase();
  let current = rootResponse.value.find(f => f.displayName.toLowerCase() === rootName);
  if (!current) {
    console.error(`Root segment "${segments[0]}" not found among top-level folders`);
    return null;
  }

  // Walk remaining segments
  for (let i = 1; i < segments.length; i++) {
    if (!current || current.childFolderCount === 0) {
      console.error(`Segment "${segments[i - 1]}" has no children; cannot resolve "${segments[i]}"`);
      return null;
    }
    const childResponse = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${current.id}/childFolders`,
      null,
      { $top: 100, $select: 'id,displayName,childFolderCount' }
    );
    if (!childResponse.value) {
      return null;
    }
    const next = segments[i].toLowerCase();
    current = childResponse.value.find(f => f.displayName.toLowerCase() === next);
    if (!current) {
      console.error(`Segment "${segments[i]}" not found under "${segments[i - 1]}"`);
      return null;
    }
  }

  console.error(`Resolved path "${segments.join('/')}" to ID: ${current.id}`);
  return current.id;
}

/**
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $top: 100,
        $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
      }
    );
    
    if (!response.value) {
      return [];
    }
    
    // Get child folders for folders with children
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);
    
    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          { 
            $select: 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount'
          }
        );
        
        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });
    
    const childFolders = await Promise.all(childFolderPromises);
    
    // Combine top-level folders and all child folders
    return [...response.value, ...childFolders.flat()];
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};
