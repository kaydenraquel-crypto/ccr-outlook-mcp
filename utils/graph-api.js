/**
 * Microsoft Graph API helpers.
 *
 * App-only mode: every `me/...` path written by callers is rewritten to
 * `users/{mailbox}/...` using the mailbox in the async context.
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');
const { getMailbox } = require('./context');

function rewritePathForMailbox(path) {
  if (!path || typeof path !== 'string') return path;
  if (path.startsWith('http://') || path.startsWith('https://')) return path;
  const mailbox = getMailbox();
  if (!mailbox) return path;

  // `me` exactly, or `me/...`
  if (path === 'me') return `users/${encodeURIComponent(mailbox)}`;
  if (path.startsWith('me/')) {
    return `users/${encodeURIComponent(mailbox)}/${path.slice(3)}`;
  }
  return path;
}

async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}) {
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  const rewrittenPath = rewritePathForMailbox(path);

  try {
    console.error(`Graph call: ${method} ${rewrittenPath}`);

    let finalUrl;
    if (rewrittenPath.startsWith('http://') || rewrittenPath.startsWith('https://')) {
      finalUrl = rewrittenPath;
    } else {
      // Don't re-encode path — mailbox is already encoded, and other segments
      // (ids, well-known folder names) are either already URL-safe or encoded at call site.
      let queryString = '';
      if (queryParams && Object.keys(queryParams).length > 0) {
        const qp = { ...queryParams };
        const filter = qp.$filter;
        delete qp.$filter;

        const params = new URLSearchParams();
        for (const [key, value] of Object.entries(qp)) {
          params.append(key, value);
        }
        queryString = params.toString();

        if (filter) {
          queryString += (queryString ? '&' : '') + `$filter=${encodeURIComponent(filter)}`;
        }
        if (queryString) queryString = '?' + queryString;
      }
      finalUrl = `${config.GRAPH_API_ENDPOINT}${rewrittenPath}${queryString}`;
    }

    return new Promise((resolve, reject) => {
      const req = https.request(
        finalUrl,
        {
          method,
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
        },
        (res) => {
          let responseData = '';
          res.on('data', (chunk) => (responseData += chunk));
          res.on('end', () => {
            if (res.statusCode >= 200 && res.statusCode < 300) {
              try {
                resolve(responseData ? JSON.parse(responseData) : {});
              } catch (e) {
                reject(new Error(`Error parsing API response: ${e.message}`));
              }
            } else if (res.statusCode === 401) {
              reject(new Error('UNAUTHORIZED'));
            } else {
              reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
            }
          });
        }
      );
      req.on('error', (err) => reject(new Error(`Network error during API call: ${err.message}`)));
      if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
        req.write(JSON.stringify(data));
      }
      req.end();
    });
  } catch (error) {
    console.error('Error calling Graph API:', error);
    throw error;
  }
}

async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') throw new Error('Pagination only supports GET requests');

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;

  do {
    const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);
    if (response.value && Array.isArray(response.value)) {
      allItems.push(...response.value);
    }
    if (maxCount > 0 && allItems.length >= maxCount) break;

    nextLink = response['@odata.nextLink'];
    if (nextLink) {
      currentUrl = nextLink;
      currentParams = {};
    }
  } while (nextLink);

  const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;
  return { value: finalItems, '@odata.count': finalItems.length };
}

module.exports = { callGraphAPI, callGraphAPIPaginated };
