#!/usr/bin/env node
/**
 * Account Switcher Server
 * Provides a simple web interface to switch between accounts
 * Useful for creating browser bookmarks to quickly switch accounts
 *
 * Usage: node account-switcher-server.js [port]
 * Default port: 3334
 *
 * Then create bookmarks like:
 *   http://localhost:3334/switch/pgcurtis90@hotmail.com
 */

const http = require('http');
const url = require('url');
const { exec } = require('child_process');
const TokenStorage = require('./auth/token-storage');
const config = require('./config');

const tokenStorage = new TokenStorage(config.AUTH_CONFIG);
const PORT = process.argv[2] || 3334;

// Helper to run shell commands
function runCommand(command) {
  return new Promise((resolve, reject) => {
    exec(command, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      } else {
        resolve(stdout);
      }
    });
  });
}

// HTML Templates
function getPageHTML(content) {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Outlook Account Switcher</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .container {
      background: white;
      border-radius: 12px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      padding: 40px;
      max-width: 600px;
      width: 100%;
    }
    h1 {
      color: #333;
      margin-bottom: 10px;
      font-size: 28px;
    }
    .subtitle {
      color: #666;
      margin-bottom: 30px;
      font-size: 14px;
    }
    .accounts-grid {
      display: grid;
      gap: 15px;
      margin-bottom: 30px;
    }
    .account-btn {
      display: block;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      text-decoration: none;
      border-radius: 8px;
      text-align: center;
      font-size: 16px;
      font-weight: 500;
      transition: transform 0.2s, box-shadow 0.2s;
      border: none;
      cursor: pointer;
    }
    .account-btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
    }
    .account-btn.active {
      background: linear-gradient(135deg, #00d4ff 0%, #0099ff 100%);
      box-shadow: 0 0 0 3px rgba(0, 212, 255, 0.3);
    }
    .status {
      padding: 15px;
      border-radius: 8px;
      margin-bottom: 20px;
      font-size: 14px;
    }
    .status.success {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
    }
    .status.error {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
    }
    .status.info {
      background: #d1ecf1;
      color: #0c5460;
      border: 1px solid #bee5eb;
    }
    .instructions {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      font-size: 13px;
      color: #666;
      line-height: 1.6;
    }
    .instructions h3 {
      color: #333;
      margin-bottom: 10px;
      font-size: 14px;
    }
    .instructions ol {
      margin-left: 20px;
    }
    .instructions li {
      margin-bottom: 8px;
    }
    .bookmark-code {
      background: #2d2d2d;
      color: #f8f8f2;
      padding: 10px;
      border-radius: 4px;
      font-family: 'Monaco', 'Menlo', monospace;
      font-size: 12px;
      overflow-x: auto;
      margin-top: 10px;
    }
    .footer {
      text-align: center;
      margin-top: 30px;
      font-size: 12px;
      color: #999;
    }
  </style>
</head>
<body>
  <div class="container">
    ${content}
  </div>
</body>
</html>
  `;
}

async function handleRequest(req, res) {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  try {
    // Main page
    if (pathname === '/') {
      const accounts = await tokenStorage.getAllAccounts();
      const activeAccount = await tokenStorage.getActiveAccount();

      let accountsHTML = '<div class="accounts-grid">';
      accounts.forEach(account => {
        const isActive = account === activeAccount;
        const className = isActive ? 'account-btn active' : 'account-btn';
        accountsHTML += `<a href="/switch/${encodeURIComponent(account)}" class="${className}">
          ${isActive ? '✓ ' : ''}${account}
        </a>`;
      });
      accountsHTML += '</div>';

      const html = getPageHTML(`
        <h1>📧 Outlook Account Switcher</h1>
        <p class="subtitle">Click an account to switch to it</p>

        ${accountsHTML}

        <div class="instructions">
          <h3>📌 To Create Browser Bookmarks:</h3>
          <ol>
            <li>Right-click the account button above</li>
            <li>Select "Bookmark Link"</li>
            <li>Save it in your Bookmarks Bar</li>
            <li>When you click the bookmark:
              <ul style="margin-left: 20px; margin-top: 8px;">
                <li>The account will be switched</li>
                <li>The service will restart automatically</li>
                <li>Then restart Claude Desktop</li>
              </ul>
            </li>
          </ol>
        </div>

        <div class="footer">
          <p>Server running on port ${PORT}</p>
          <p>Keep this server running while using account switching</p>
        </div>
      `);

      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
    }
    // Switch account endpoint
    else if (pathname.startsWith('/switch/')) {
      const email = decodeURIComponent(pathname.replace('/switch/', ''));
      const accounts = await tokenStorage.getAllAccounts();

      if (!accounts.includes(email)) {
        const html = getPageHTML(`
          <h1>❌ Error</h1>
          <div class="status error">Account "${email}" not found</div>
          <a href="/" class="account-btn" style="display: inline-block; width: auto; padding: 10px 20px;">← Back</a>
        `);
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(html);
        return;
      }

      // Switch account
      await tokenStorage.setActiveAccount(email);

      // Restart the launchd service
      try {
        await runCommand('launchctl stop com.outlook-mcp.server');
        await new Promise(resolve => setTimeout(resolve, 1000));
        await runCommand('launchctl start com.outlook-mcp.server');
      } catch (error) {
        console.error('Error restarting service:', error);
      }

      const html = getPageHTML(`
        <h1>✅ Account Switched!</h1>
        <div class="status success">
          <strong>Active account:</strong> ${email}<br>
          <strong>Service:</strong> Restarted automatically
        </div>

        <div class="instructions">
          <h3>📌 Next Steps:</h3>
          <ol>
            <li><strong>Close Claude Desktop</strong> (Command+Q)</li>
            <li><strong>Wait 2 seconds</strong></li>
            <li><strong>Reopen Claude Desktop</strong></li>
            <li>You're now connected to <strong>${email}</strong></li>
          </ol>
        </div>

        <a href="/" class="account-btn" style="display: inline-block; width: auto; padding: 10px 20px;">← Back</a>

        <div class="footer">
          <p style="margin-top: 20px;">💡 Pro tip: You can bookmark this account for quick access!</p>
        </div>
      `);

      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
    }
    // 404
    else {
      res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(getPageHTML(`
        <h1>404 - Not Found</h1>
        <p>Page not found</p>
        <a href="/" class="account-btn" style="display: inline-block; width: auto; padding: 10px 20px;">← Back Home</a>
      `));
    }
  } catch (error) {
    console.error('Error:', error);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(getPageHTML(`
      <h1>❌ Error</h1>
      <div class="status error">
        <strong>Error:</strong> ${error.message}
      </div>
      <a href="/" class="account-btn" style="display: inline-block; width: auto; padding: 10px 20px;">← Back</a>
    `));
  }
}

// Start server
const server = http.createServer(handleRequest);
server.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════╗
║    Outlook Account Switcher Server                          ║
║                                                              ║
║    📱 Open in browser: http://localhost:${PORT}               ║
║                                                              ║
║    💡 Bookmark these URLs in your browser:                  ║
║    ${`http://localhost:${PORT}/switch/[email]`.padEnd(56)}║
║                                                              ║
║    🎯 Click a bookmark to:                                  ║
║      1. Switch to that account                              ║
║      2. Restart the MCP service                             ║
║      3. Then restart Claude Desktop                         ║
║                                                              ║
║    Press Ctrl+C to stop                                     ║
╚════════════════════════════════════════════════════════════╝
  `);
});

process.on('SIGINT', () => {
  console.log('\nShutting down account switcher server...');
  server.close();
  process.exit(0);
});
