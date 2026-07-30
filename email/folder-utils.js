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
 * Resolve a single folder segment within a parent folder
 * @param {string} accessToken - Access token
 * @param {string|null} parentId - Parent folder ID, or null for top-level
 * @param {string} segment - Folder segment name to resolve
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function resolveSegmentInParent(accessToken, parentId, segment) {
  const base = parentId ? `me/mailFolders/${parentId}/childFolders` : 'me/mailFolders';

  // Escape single quotes for OData string literal (apostrophes must be doubled)
  const escapedSegment = segment.replace(/'/g, "''");

  // First try with exact match filter
  const response = await callGraphAPI(
    accessToken,
    'GET',
    base,
    null,
    { $filter: `displayName eq '${escapedSegment}'` }
  );

  if (response.value && response.value.length > 0) {
    return response.value[0].id;
  }

  // If exact match fails, try to get all folders and do a case-insensitive comparison
  const allFoldersResponse = await callGraphAPI(
    accessToken,
    'GET',
    base,
    null,
    { $top: 100 }
  );

  if (allFoldersResponse.value) {
    const lowerSegment = segment.toLowerCase();
    const matchingFolder = allFoldersResponse.value.find(
      folder => folder.displayName.toLowerCase() === lowerSegment
    );

    if (matchingFolder) {
      return matchingFolder.id;
    }
  }

  return null;
}

/**
 * Get the ID of a mail folder by its name or path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name or path (e.g. "Tramite/REQ-104951") of the folder to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  const segments = folderName.split('/').map(s => s.trim()).filter(Boolean);
  if (segments.length === 0) {
    return null;
  }

  try {
    let currentId = null;
    for (const segment of segments) {
      currentId = await resolveSegmentInParent(accessToken, currentId, segment);
      if (currentId === null) {
        return null;
      }
    }
    return currentId;
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
  resolveSegmentInParent,
  getFolderIdByName,
  getAllFolders
};
