/**
 * Power Automate / Flow API client
 * Separate from Graph API as it uses different endpoint and auth scope
 */
const https = require('https');
const config = require('../config');

/**
 * Makes a request to the Power Automate Flow API
 * @param {string} accessToken - The Flow access token
 * @param {string} method - HTTP method
 * @param {string} path - API endpoint path (after /providers/Microsoft.ProcessSimple)
 * @param {object} data - Data to send for POST/PUT requests
 * @returns {Promise<object>} - The API response
 */
async function callFlowAPI(accessToken, method, path, data = null) {
  // For test tokens, simulate response
  if (config.USE_TEST_MODE && accessToken.startsWith('test_')) {
    console.error(`TEST MODE: Simulating Flow ${method} ${path}`);
    return simulateFlowResponse(method, path);
  }

  return new Promise((resolve, reject) => {
    // Flow API uses a different URL structure
    const fullUrl = `${config.FLOW_API_ENDPOINT}/providers/Microsoft.ProcessSimple${path}?api-version=2016-11-01`;
    console.error(`Making Flow API call: ${method} ${fullUrl}`);

    const urlObj = new URL(fullUrl);

    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: method,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(options, (res) => {
      let responseData = '';

      res.on('data', (chunk) => {
        responseData += chunk;
      });

      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(responseData || '{}'));
          } catch (error) {
            reject(new Error(`Error parsing Flow API response: ${error.message}`));
          }
        } else if (res.statusCode === 401) {
          reject(new Error('FLOW_UNAUTHORIZED'));
        } else if (res.statusCode === 403) {
          reject(new Error('Access denied. Ensure your account has Power Automate access and the flow is solution-aware.'));
        } else {
          reject(new Error(`Flow API call failed with status ${res.statusCode}: ${responseData}`));
        }
      });
    });

    req.on('error', (error) => {
      reject(new Error(`Network error during Flow API call: ${error.message}`));
    });

    if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
      req.write(JSON.stringify(data));
    }

    req.end();
  });
}

/**
 * Simulate Flow API response for test mode
 */
function simulateFlowResponse(method, path) {
  if (path.includes('/environments')) {
    return {
      value: [
        {
          name: 'Default-12345',
          properties: {
            displayName: 'Default Environment',
            isDefault: true
          }
        }
      ]
    };
  }

  if (path.includes('/flows') && !path.includes('/runs')) {
    return {
      value: [
        {
          name: 'flow-123',
          properties: {
            displayName: 'Test Flow',
            state: 'Started',
            createdTime: new Date().toISOString()
          }
        }
      ]
    };
  }

  if (path.includes('/runs')) {
    return {
      value: [
        {
          name: 'run-123',
          properties: {
            status: 'Succeeded',
            startTime: new Date().toISOString()
          }
        }
      ]
    };
  }

  return {};
}

module.exports = {
  callFlowAPI
};
