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
  // Map well-known folder names to their Graph API aliases
  const WELL_KNOWN_FOLDER_IDS = {
    'deleted items': 'deleteditems',
    'deleted': 'deleteditems',
    'sent items': 'sentitems',
    'sent': 'sentitems',
    'drafts': 'drafts',
    'inbox': 'inbox',
    'junk email': 'junkemail',
    'junk': 'junkemail',
    'archive': 'archive'
  };

  try {
    // Check well-known folders first — these have fixed IDs in Graph API
    const wellKnownId = WELL_KNOWN_FOLDER_IDS[folderName.toLowerCase()];
    if (wellKnownId) {
      console.error(`Using well-known folder ID "${wellKnownId}" for "${folderName}"`);
      const wkResponse = await callGraphAPI(accessToken, 'GET', `me/mailFolders/${wellKnownId}`);
      if (wkResponse && wkResponse.id) {
        console.error(`Resolved well-known folder "${folderName}" to ID: ${wkResponse.id}`);
        return wkResponse.id;
      }
    }

    // Try exact match filter
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
    
    // If exact match fails, try all folders (including child folders) with case-insensitive comparison
    console.error(`No exact match found for "${folderName}", trying full folder search`);
    const allFolders = await getAllFolders(accessToken);

    if (allFolders.length > 0) {
      const lowerFolderName = folderName.toLowerCase();
      const matchingFolder = allFolders.find(
        folder => folder.displayName.toLowerCase() === lowerFolderName
      );

      if (matchingFolder) {
        console.error(`Found match for "${folderName}" with ID: ${matchingFolder.id}`);
        return matchingFolder.id;
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
 * Get all mail folders
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  const selectFields = 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount';
  try {
    return await fetchFoldersRecursive(accessToken, 'me/mailFolders', selectFields);
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

/**
 * Recursively fetch folders and all their descendants
 * @param {string} accessToken - Access token
 * @param {string} endpoint - Graph API endpoint to fetch folders from
 * @param {string} selectFields - Fields to select
 * @returns {Promise<Array>} - Flat array of all folder objects
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

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};
