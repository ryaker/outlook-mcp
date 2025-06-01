#!/usr/bin/env node
/**
 * Backwards-compatible Outlook MCP Server with integrated OAuth authentication
 * Supports both Streamable HTTP and legacy SSE transports
 */

const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const { SSEServerTransport } = require("@modelcontextprotocol/sdk/server/sse.js");
const { isInitializeRequest } = require("@modelcontextprotocol/sdk/types.js");
const express = require('express');
const { randomUUID } = require("node:crypto");

const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const config = require('./config');
const TokenStorage = require('./auth/token-storage');
const { setupOAuthRoutes, createAuthConfig } = require('./auth/oauth-server');
// Load environment variables
require('dotenv').config();

// Import module tools
const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');

// Initialize token storage
const tokenStorage = new TokenStorage();

// Authentication configuration  
const PORT = process.env.SSE_PORT ? parseInt(process.env.SSE_PORT, 10) : 3333;
const AUTH_CONFIG = createAuthConfig(PORT);

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} SSE MCP SERVER`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// Combine all tools
const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools
];

// Store transports for each session type
const transports = {
  streamable: {},
  sse: {}
};

/**
 * Create MCP server instance
 */
function createMCPServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    {
      capabilities: {
        tools: TOOLS.reduce((acc, tool) => {
          acc[tool.name] = {};
          return acc;
        }, {})
      }
    }
  );

  // Handle all MCP requests
  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      if (method === "initialize") {
        console.error(`INITIALIZE REQUEST: ID [${id}]`);
        return {
          protocolVersion: "2024-11-05",
          capabilities: {
            tools: TOOLS.reduce((acc, tool) => {
              acc[tool.name] = {};
              return acc;
            }, {})
          },
          serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
        };
      }

      if (method === "tools/list") {
        return {
          tools: TOOLS.map(tool => ({
            name: tool.name,
            description: tool.description,
            inputSchema: tool.inputSchema
          }))
        };
      }

      if (method === "resources/list") return { resources: [] };
      if (method === "prompts/list") return { prompts: [] };

      if (method === "tools/call") {
        const { name, arguments: args = {} } = params || {};
        const tool = TOOLS.find(t => t.name === name);

        if (tool && tool.handler) {
          return await tool.handler(args);
        }

        return {
          error: {
            code: -32601,
            message: `Tool not found: ${name}`
          }
        };
      }

      return {
        error: {
          code: -32601,
          message: `Method not found: ${method}`
        }
      };
    } catch (error) {
      console.error(`Error in MCP request handler:`, error);
      return {
        error: {
          code: -32603,
          message: `Error processing request: ${error.message}`
        }
      };
    }
  };

  return server;
}
const app = express();
app.use(express.json());

// Modern Streamable HTTP endpoint
app.all('/mcp', async (req, res) => {
  // Check for existing session ID
  const sessionId = req.headers['mcp-session-id'];
  let transport;

  if (sessionId && transports.streamable[sessionId]) {
    // Reuse existing transport
    transport = transports.streamable[sessionId];
  } else if (!sessionId && isInitializeRequest(req.body)) {
    // New initialization request
    transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => randomUUID(),
      onsessioninitialized: (sessionId) => {
        // Store the transport by session ID
        transports.streamable[sessionId] = transport;
      }
    });

    // Clean up transport when closed
    transport.onclose = () => {
      if (transport.sessionId) {
        delete transports.streamable[transport.sessionId];
      }
    };
    const server = createMCPServer();
    await server.connect(transport);
  } else {
    res.status(400).json({
      jsonrpc: '2.0',
      error: {
        code: -32000,
        message: 'Bad Request: No valid session ID provided',
      },
      id: null,
    });
    return;
  }

  await transport.handleRequest(req, res, req.body);
});

// Legacy SSE endpoint for older clients
app.get('/sse', async (req, res) => {
  console.error('Legacy SSE connection requested');

  const transport = new SSEServerTransport('/messages', res);
  transports.sse[transport.sessionId] = transport;

  res.on("close", () => {
    console.error(`Legacy SSE connection closed: ${transport.sessionId}`);
    delete transports.sse[transport.sessionId];
  });

  const server = createMCPServer();
  await server.connect(transport);
});

// Legacy message endpoint for older clients
app.post('/messages', async (req, res) => {
  const sessionId = req.query.sessionId;
  const transport = transports.sse[sessionId];

  if (!transport) {
    res.status(400).send('No SSE transport found for sessionId');
    return;
  }

  await transport.handlePostMessage(req, res, req.body);
});

// Setup OAuth routes using shared module
setupOAuthRoutes(app, AUTH_CONFIG, tokenStorage);

// Root path - server info
app.get('/', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Outlook MCP Server</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
          h1 { color: #0078d4; }
          .info-box { background-color: #e7f6fd; border: 1px solid #b3e0ff; padding: 15px; border-radius: 4px; margin: 10px 0; }
          .endpoint { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; font-family: monospace; }
        </style>
      </head>
      <body>
        <h1>Outlook MCP Server</h1>
        <div class="info-box">
          <p>This server supports both modern Streamable HTTP and legacy SSE transports.</p>
          <p><strong>Available Endpoints:</strong></p>
          <ul>
            <li><span class="endpoint">/mcp</span> - Modern Streamable HTTP endpoint (recommended)</li>
            <li><span class="endpoint">/sse</span> - Legacy SSE endpoint</li>
            <li><span class="endpoint">/messages</span> - Legacy SSE message endpoint</li>
            <li><span class="endpoint">/auth</span> - Start OAuth authentication</li>
            <li><span class="endpoint">/auth/callback</span> - OAuth callback handler</li>
            <li><span class="endpoint">/token-status</span> - Check token status</li>
          </ul>
        </div>
        <div class="info-box">
          <p><strong>Features:</strong></p>
          <ul>
            <li>Automatic token refresh before expiration</li>
            <li>HTTP streaming-based MCP communication</li>
            <li>Integrated OAuth 2.0 flow</li>
            <li>Comprehensive token management</li>
          </ul>
        </div>
        <p>Server is running at <strong>http://localhost:3333</strong></p>
      </body>
    </html>
  `);
});

// Start the server
const HOST = 'localhost';

app.listen(PORT, HOST, (error) => {
  if (error) {
    console.error(`Error starting server: ${error}`);
    process.exit(1);
  }

  console.error(`MCP Server running at http://${HOST}:${PORT}`);
  console.error(`Modern Streamable HTTP endpoint: http://${HOST}:${PORT}/mcp`);
  console.error(`Legacy SSE endpoint: http://${HOST}:${PORT}/sse`);
  console.error(`OAuth callback: ${AUTH_CONFIG.redirectUri}`);
  console.error(`Token storage: ${tokenStorage.tokenPath}`);

  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.error('\n⚠️  WARNING: Microsoft Graph API credentials are not set.');
    console.error('   Please set the MS_CLIENT_ID and MS_CLIENT_SECRET environment variables.');
  }
});

// Handle termination
function cleanup() {
  console.error('SSE MCP Server shutting down');
  process.exit(0);
}

process.on('SIGINT', cleanup);
process.on('SIGTERM', cleanup);
