/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

/**
 * Makes a request to the Microsoft Graph API
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {string} path - API endpoint path
 * @param {object} data - Data to send for POST/PUT requests
 * @param {object} queryParams - Query parameters
 * @returns {Promise<object>} - The API response
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}) {
  // For test tokens, we'll simulate the API call
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  try {
    console.error(`Making real API call: ${method} ${path}`);
    
    // Check if path already contains the full URL (from nextLink)
    let finalUrl;
    if (path.startsWith('http://') || path.startsWith('https://')) {
      // Path is already a full URL (from pagination nextLink)
      finalUrl = path;
      console.error(`Using full URL from nextLink: ${finalUrl}`);
    } else {
      // Build URL from path and queryParams
      // Split path and any existing query string to avoid double-encoding
      const [rawPath, rawQuery] = path.split('?');
      const encodedPath = rawPath.split('/')
        .map(segment => encodeURIComponent(segment))
        .join('/');

      const params = new URLSearchParams();
      let filter;
      let filterPreEncoded = false;

      // Process existing query string from path
      if (rawQuery) {
        const existingParams = new URLSearchParams(rawQuery);
        for (const [key, value] of existingParams.entries()) {
          if (key === '$filter') {
            filter = value;
            filterPreEncoded = true;
          } else {
            params.append(key, value);
          }
        }
      }

      // Process queryParams object (these override path params)
      if (queryParams && Object.keys(queryParams).length > 0) {
        const { $filter, ...rest } = queryParams;
        if ($filter !== undefined) {
          filter = $filter;
          filterPreEncoded = false;
        }
        for (const [key, value] of Object.entries(rest)) {
          params.append(key, value);
        }
      }

      // Build final query string
      let queryString = params.toString();
      if (queryString) {
        queryString = `?${queryString}`;
      }

      // Add filter parameter with proper encoding
      if (filter !== undefined) {
        const separator = queryString ? '&' : '?';
        const encodedFilter = filterPreEncoded ? filter : encodeURIComponent(filter);
        queryString += `${separator}$filter=${encodedFilter}`;
      }

      console.error(`Query string: ${queryString}`);
      finalUrl = `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
      console.error(`Full URL: ${finalUrl}`);
    }
    
    return new Promise((resolve, reject) => {
      const options = {
        method: method,
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      };
      
      const req = https.request(finalUrl, options, (res) => {
        let responseData = '';
        
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        
        res.on('end', () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            try {
              responseData = responseData ? responseData : '{}';
              const jsonResponse = JSON.parse(responseData);
              resolve(jsonResponse);
            } catch (error) {
              reject(new Error(`Error parsing API response: ${error.message}`));
            }
          } else if (res.statusCode === 401) {
            // Token expired or invalid
            reject(new Error('UNAUTHORIZED'));
          } else {
            reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
          }
        });
      });
      
      req.on('error', (error) => {
        reject(new Error(`Network error during API call: ${error.message}`));
      });
      
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

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET only for pagination)
 * @param {string} path - API endpoint path
 * @param {object} queryParams - Initial query parameters
 * @param {number} maxCount - Maximum number of items to retrieve (0 = all)
 * @returns {Promise<object>} - Combined API response with all items
 */
async function callGraphAPIPaginated(accessToken, method, path, queryParams = {}, maxCount = 0) {
  if (method !== 'GET') {
    throw new Error('Pagination only supports GET requests');
  }

  const allItems = [];
  let nextLink = null;
  let currentUrl = path;
  let currentParams = queryParams;
  let pageCount = 0;

  try {
    do {
      pageCount++;

      // Safeguard against infinite loops
      if (pageCount > config.MAX_PAGINATION_PAGES) {
        console.error(`Pagination: Reached max page limit of ${config.MAX_PAGINATION_PAGES}, stopping.`);
        break;
      }

      // Make API call
      const response = await callGraphAPI(accessToken, method, currentUrl, null, currentParams);

      // Add items from this page
      if (response.value && Array.isArray(response.value)) {
        allItems.push(...response.value);
        console.error(`Pagination: Retrieved ${response.value.length} items on page ${pageCount}, total so far: ${allItems.length}`);
      }

      // Check if we've reached the desired count
      if (maxCount > 0 && allItems.length >= maxCount) {
        console.error(`Pagination: Reached max count of ${maxCount}, stopping`);
        break;
      }

      // Get next page URL
      nextLink = response['@odata.nextLink'];

      if (nextLink) {
        // Pass the full nextLink URL directly to callGraphAPI
        currentUrl = nextLink;
        currentParams = {}; // nextLink already contains all params
        console.error(`Pagination: Following nextLink (page ${pageCount}), ${allItems.length} items so far`);
      }
    } while (nextLink);

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;

    console.error(`Pagination complete: Retrieved ${finalItems.length} total items`);
    
    return {
      value: finalItems,
      '@odata.count': finalItems.length
    };
  } catch (error) {
    console.error('Error during pagination:', error);
    throw error;
  }
}

module.exports = {
  callGraphAPI,
  callGraphAPIPaginated
};
