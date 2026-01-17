/**
 * OneDrive chunked upload functionality (files > 4MB)
 */
const https = require('https');
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const CHUNK_SIZE = 320 * 1024 * 10; // 3.2MB chunks (must be multiple of 320KB)

/**
 * Large file upload handler (chunked upload for files > 4MB)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUploadLarge(args) {
  const path = args.path;
  const content = args.content;
  const conflictBehavior = args.conflictBehavior || 'rename';

  if (!path) {
    return {
      content: [{
        type: "text",
        text: "Path is required (e.g., '/Documents/largefile.zip')."
      }]
    };
  }

  if (!content) {
    return {
      content: [{
        type: "text",
        text: "Content is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const contentBuffer = Buffer.from(content);
    const fileSize = contentBuffer.length;

    // Normalize path
    const normalizedPath = path.replace(/^\/+|\/+$/g, '');

    // Step 1: Create upload session
    const sessionEndpoint = `me/drive/root:/${normalizedPath}:/createUploadSession`;
    const sessionBody = {
      item: {
        '@microsoft.graph.conflictBehavior': conflictBehavior
      }
    };

    const sessionResponse = await callGraphAPI(accessToken, 'POST', sessionEndpoint, sessionBody);

    if (!sessionResponse || !sessionResponse.uploadUrl) {
      return {
        content: [{
          type: "text",
          text: "Failed to create upload session."
        }]
      };
    }

    const uploadUrl = sessionResponse.uploadUrl;

    // Step 2: Upload chunks
    let offset = 0;
    let response;

    while (offset < fileSize) {
      const chunkEnd = Math.min(offset + CHUNK_SIZE, fileSize);
      const chunk = contentBuffer.slice(offset, chunkEnd);

      response = await uploadChunk(uploadUrl, chunk, offset, chunkEnd - 1, fileSize);

      if (response.error) {
        return {
          content: [{
            type: "text",
            text: `Upload failed at byte ${offset}: ${response.error}`
          }]
        };
      }

      offset = chunkEnd;

      // Log progress
      const progress = Math.round((offset / fileSize) * 100);
      console.error(`Upload progress: ${progress}%`);
    }

    // Final response should contain the uploaded file info
    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Upload completed but no file info returned."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully uploaded "${response.name}" (${formatSize(response.size)})\n\nID: ${response.id}\nWeb URL: ${response.webUrl}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error uploading large file: ${error.message}`
      }]
    };
  }
}

/**
 * Upload a single chunk to the upload session
 */
async function uploadChunk(uploadUrl, chunk, start, end, totalSize) {
  return new Promise((resolve, reject) => {
    const options = {
      method: 'PUT',
      headers: {
        'Content-Length': chunk.length,
        'Content-Range': `bytes ${start}-${end}/${totalSize}`
      }
    };

    const req = https.request(uploadUrl, options, (res) => {
      let responseData = '';

      res.on('data', (data) => {
        responseData += data;
      });

      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(responseData || '{}'));
          } catch (e) {
            resolve({});
          }
        } else {
          resolve({ error: `Status ${res.statusCode}: ${responseData}` });
        }
      });
    });

    req.on('error', (error) => {
      resolve({ error: error.message });
    });

    req.write(chunk);
    req.end();
  });
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = handleUploadLarge;
