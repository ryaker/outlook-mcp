#!/usr/bin/env node
const express = require('express');
const TokenStorage = require('./auth/token-storage');
const { setupOAuthRoutes, createAuthConfig } = require('./auth/oauth-server');

// Load environment variables from .env file
require('dotenv').config();

// Log to console
console.log('Starting Outlook Authentication Server');

// Initialize token storage and auth config
const tokenStorage = new TokenStorage();
const AUTH_CONFIG = createAuthConfig(3333);

// Create Express app
const app = express();
app.use(express.json());

// Setup OAuth routes using shared module
setupOAuthRoutes(app, AUTH_CONFIG, tokenStorage);

// Root path - provide instructions
app.get('/', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Outlook Authentication Server</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
          h1 { color: #0078d4; }
          .info-box { background-color: #e7f6fd; border: 1px solid #b3e0ff; padding: 15px; border-radius: 4px; }
          code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h1>Outlook Authentication Server</h1>
        <div class="info-box">
          <p>This server is running to handle Microsoft Graph API authentication callbacks.</p>
          <p>Don't navigate here directly. Instead, use the <code>authenticate</code> tool in Claude to start the authentication process.</p>
          <p>Make sure you've set the <code>MS_CLIENT_ID</code> and <code>MS_CLIENT_SECRET</code> environment variables.</p>
        </div>
        <p>Server is running at http://localhost:3333</p>
      </body>
    </html>
  `);
});


// Start server
const PORT = 3333;
app.listen(PORT, () => {
  console.log(`Authentication server running at http://localhost:${PORT}`);
  console.log(`Waiting for authentication callback at ${AUTH_CONFIG.redirectUri}`);
  console.log(`Token will be stored at: ${tokenStorage.tokenPath}`);
  
  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log('\n⚠️  WARNING: Microsoft Graph API credentials are not set.');
    console.log('   Please set the MS_CLIENT_ID and MS_CLIENT_SECRET environment variables.');
  }
});

// Handle termination
process.on('SIGINT', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});
