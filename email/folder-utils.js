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
 * @param {object} [options]
 * @param {boolean} [options.strict=false] - Throw when the folder cannot be found
 *   instead of silently falling back to the inbox
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName, options = {}) {
  const { strict = false } = options;

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

  let folderId = null;
  try {
    // Try to find the folder by name, including nested subfolders
    folderId = await getFolderIdByName(accessToken, folderName, { deep: true });
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    if (strict) {
      throw error;
    }
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  if (folderId) {
    const path = `me/mailFolders/${folderId}/messages`;
    console.error(`Resolved folder "${folderName}" to path: ${path}`);
    return path;
  }

  if (strict) {
    throw new Error(`Mail folder "${folderName}" not found`);
  }

  // If not found, fall back to inbox
  console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
  return WELL_KNOWN_FOLDERS['inbox'];
}

/**
 * Get the ID of a mail folder by its name
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @param {object} [options]
 * @param {boolean} [options.deep=false] - Also search nested subfolders
 *   (breadth-first) when no top-level folder matches
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName, options = {}) {
  const { deep = false } = options;
  try {
    // First try with exact match filter
    console.error(`Looking for folder with name "${folderName}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${folderName}'` }
    );
    
    if (response.value && response.value.length > 0) {
      console.error(`Found folder "${folderName}" with ID: ${response.value[0].id}`);
      return response.value[0].id;
    }
    
    // If exact match fails, try to get all folders and do a case-insensitive comparison
    console.error(`No exact match found for "${folderName}", trying case-insensitive search`);
    const allFoldersResponse = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $top: 100 }
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

      if (deep) {
        const nestedId = await findNestedFolderId(accessToken, allFoldersResponse.value, lowerFolderName);
        if (nestedId) {
          console.error(`Found nested folder "${folderName}" with ID: ${nestedId}`);
          return nestedId;
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

/** Maximum folder depth searched below the top level */
const MAX_FOLDER_DEPTH = 6;

/**
 * Breadth-first search through child folders for a case-insensitive name match
 * @param {string} accessToken - Access token
 * @param {Array} topLevelFolders - Already fetched top-level folders
 * @param {string} lowerFolderName - Lower-cased folder name to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function findNestedFolderId(accessToken, topLevelFolders, lowerFolderName) {
  let level = topLevelFolders.filter(f => f.childFolderCount > 0);

  for (let depth = 0; depth < MAX_FOLDER_DEPTH && level.length > 0; depth++) {
    const children = (await Promise.all(level.map(async (folder) => {
      try {
        const response = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          { $top: 100, $select: 'id,displayName,childFolderCount' }
        );
        return (response && response.value) || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    }))).flat();

    const match = children.find(f => f.displayName && f.displayName.toLowerCase() === lowerFolderName);
    if (match) {
      return match.id;
    }

    level = children.filter(f => f.childFolderCount > 0);
  }

  return null;
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
