/**
 * Authentication facade — app-only (client credentials) via MSAL.
 * The legacy `authenticate` tool is not needed in app-only mode; tools array is empty.
 */
const tokenManager = require('./token-manager');

async function ensureAuthenticated() {
  return tokenManager.getAccessToken();
}

const authTools = [];

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated,
};
