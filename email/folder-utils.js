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
 * Now searches top-level folders
 * and one level of child folders (e.g. Inbox subfolders).
 */
async function getFolderIdByName(accessToken, folderName) {
  const { callGraphAPI } = require('../utils/graph-api');
  const normalised = folderName.trim().toLowerCase();

  // 1. Fetch all top-level mail folders
  const topLevel = await callGraphAPI(accessToken, 'GET', 'me/mailFolders?$top=50');
  const folders = topLevel.value || [];

  // 2. Check top-level folders first
  for (const folder of folders) {
    if (folder.displayName.toLowerCase() === normalised) {
      return folder.id;
    }
  }

  // 3. Not found at top level — search one level of child folders
  for (const folder of folders) {
    const children = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${folder.id}/childFolders?$top=50`
    );
    for (const child of (children.value || [])) {
      if (child.displayName.toLowerCase() === normalised) {
        return child.id;
      }
    }
  }

  // 4. Not found anywhere
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
