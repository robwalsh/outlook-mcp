/**
 * Authentication module for Outlook MCP server
 */
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');

// Singleton TokenStorage instance for automatic token refresh
const tokenStorage = new TokenStorage();

/**
 * Ensures the user is authenticated and returns an access token.
 * Automatically refreshes expired tokens using the refresh_token grant.
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  // Use TokenStorage which handles automatic refresh
  const accessToken = await tokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return accessToken;
}

/**
 * Forces a refresh of the access token, regardless of its current expiry.
 * Used to recover from a 401 when the access token was invalidated server-side
 * before its expires_at (e.g. revocation, CAE, password/MFA change). Returns the
 * fresh access token, or null if no usable refresh token is available.
 * @returns {Promise<string|null>} - A new access token, or null
 */
async function refreshActiveToken() {
  try {
    await tokenStorage.getTokens(); // ensure tokens are loaded from disk
    return await tokenStorage.refreshAccessToken();
  } catch (error) {
    console.error('Forced token refresh failed:', error && error.message);
    return null;
  }
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated,
  refreshActiveToken
};
