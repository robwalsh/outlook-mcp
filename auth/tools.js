/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }
  
  // For real authentication, generate an auth URL and instruct the user to visit it
  const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  
  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft: ${authUrl}\n\nAfter authentication, you will be redirected back to this application.`
    }]
  };
}

/**
 * Check authentication status tool handler
 *
 * Answers the question a caller actually has: "will my next call work?"
 * The old implementation inspected the raw token file and lied in both
 * directions: an expired access token with a perfectly good refresh token
 * read as "Not authenticated" (access tokens expire hourly; ensureAuthenticated
 * refreshes silently, so real calls worked fine), and a server-side-revoked
 * token that had not yet passed expires_at read as "Authenticated and ready"
 * while every call failed. Now it goes through the same path real calls use
 * (refresh included) and then proves the token against Graph with /me.
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');
  // Lazy require: auth/index.js requires this file at load time.
  const { ensureAuthenticated } = require('./index');
  let accessToken;
  try {
    accessToken = await ensureAuthenticated();
  } catch (e) {
    console.error('[CHECK-AUTH-STATUS] No usable token (refresh included):', e.message);
    return {
      content: [{ type: "text", text: "Not authenticated — no usable token and no refreshable session. Run authenticate." }]
    };
  }
  try {
    const { callGraphAPI } = require('../utils/graph-api');
    const me = await callGraphAPI(accessToken, 'GET', 'me', null, { $select: 'userPrincipalName' });
    console.error(`[CHECK-AUTH-STATUS] Graph /me ok (${me && me.userPrincipalName})`);
    return {
      content: [{ type: "text", text: `Authenticated and ready (verified against Graph as ${me && me.userPrincipalName ? me.userPrincipalName : 'unknown account'})` }]
    };
  } catch (e) {
    console.error('[CHECK-AUTH-STATUS] Token held locally but Graph rejected it:', e.message);
    return {
      content: [{ type: "text", text: `Token present but Graph REJECTED it (${String(e.message).slice(0, 120)}) — the session was invalidated server-side. Run authenticate.` }]
    };
  }
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
