/**
 * Token manager — Microsoft Graph app-only auth via MSAL client credentials flow.
 * No disk persistence; MSAL keeps an in-memory cache and we just re-request on expiry.
 */
const { ConfidentialClientApplication } = require('@azure/msal-node');
const config = require('../config');

let cachedClient = null;

function getClient() {
  if (cachedClient) return cachedClient;

  const { clientId, clientSecret, tenantId } = config.AUTH_CONFIG;
  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      'Missing Microsoft Entra credentials. Set MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID.'
    );
  }

  cachedClient = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  });
  return cachedClient;
}

async function getAccessToken() {
  if (config.USE_TEST_MODE) {
    return 'test_access_token_' + Date.now();
  }
  const client = getClient();
  const result = await client.acquireTokenByClientCredential({
    scopes: config.AUTH_CONFIG.scopes,
  });
  if (!result || !result.accessToken) {
    throw new Error('Failed to acquire access token from Microsoft Entra.');
  }
  return result.accessToken;
}

module.exports = { getAccessToken };
