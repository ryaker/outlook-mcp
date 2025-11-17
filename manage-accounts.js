#!/usr/bin/env node
/**
 * Account Management Utility
 * Allows viewing and switching between multiple authenticated Outlook accounts
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');

const tokenPath = path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json');

// Helper to read tokens
function loadTokens() {
  try {
    if (!fs.existsSync(tokenPath)) {
      console.log('No tokens file found.');
      return null;
    }
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    return JSON.parse(tokenData);
  } catch (error) {
    console.error('Error reading tokens:', error.message);
    return null;
  }
}

// Helper to save tokens
function saveTokens(tokens) {
  try {
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.log('✓ Tokens saved successfully');
  } catch (error) {
    console.error('Error saving tokens:', error.message);
  }
}

// List all accounts
function listAccounts() {
  const tokens = loadTokens();
  if (!tokens) {
    console.log('No authenticated accounts found.');
    return;
  }

  // Handle new multi-account format
  if (tokens.accounts && typeof tokens.accounts === 'object') {
    const accounts = Object.keys(tokens.accounts);
    const activeAccount = tokens.activeAccount;

    console.log('\n📧 Authenticated Accounts:');
    console.log('─'.repeat(50));

    accounts.forEach((email, index) => {
      const isActive = email === activeAccount ? ' ✓ (ACTIVE)' : '';
      const token = tokens.accounts[email];
      const expiresAt = new Date(token.expires_at);
      const status = Date.now() < token.expires_at ? '✓ Valid' : '✗ Expired';

      console.log(`${index + 1}. ${email}${isActive}`);
      console.log(`   Status: ${status}`);
      console.log(`   Expires: ${expiresAt.toLocaleString()}`);
    });
    console.log('─'.repeat(50) + '\n');
    return accounts;
  }

  // Handle old single-token format
  console.log('Old token format detected. Please re-authenticate your accounts.');
  return [];
}

// Switch active account
function switchAccount(email) {
  const tokens = loadTokens();
  if (!tokens || !tokens.accounts || !tokens.accounts[email]) {
    console.error(`Error: Account '${email}' not found.`);
    return false;
  }

  tokens.activeAccount = email;
  saveTokens(tokens);
  console.log(`✓ Switched to account: ${email}`);
  return true;
}

// Clear old single tokens and prepare for new accounts
function clearAllTokens() {
  try {
    if (fs.existsSync(tokenPath)) {
      fs.unlinkSync(tokenPath);
      console.log('✓ Old tokens cleared. You can now re-authenticate.');
    }
  } catch (error) {
    console.error('Error clearing tokens:', error.message);
  }
}

// Interactive menu
function showMenu() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  const menu = `
╔════════════════════════════════════════════════════════════╗
║           Outlook MCP Account Manager                       ║
╚════════════════════════════════════════════════════════════╝

1. List all authenticated accounts
2. Switch active account
3. Re-authenticate (clear all and start fresh)
4. View current active account
5. Exit

  `;

  console.log(menu);

  rl.question('Select an option (1-5): ', (answer) => {
    switch (answer.trim()) {
      case '1':
        listAccounts();
        rl.close();
        showMenu();
        break;

      case '2':
        const accounts = listAccounts();
        if (accounts.length > 0) {
          rl.question('Enter the email address to activate: ', (email) => {
            switchAccount(email.trim());
            rl.close();
            showMenu();
          });
        } else {
          rl.close();
          showMenu();
        }
        break;

      case '3':
        rl.question('⚠️  This will clear ALL tokens. Are you sure? (yes/no): ', (confirm) => {
          if (confirm.toLowerCase() === 'yes') {
            clearAllTokens();
            console.log('\nTo re-authenticate:');
            console.log('  1. Run: npm run auth-server');
            console.log('  2. In another terminal: npx @modelcontextprotocol/inspector node index.js');
            console.log('  3. Use the authenticate tool in the inspector');
          }
          rl.close();
          showMenu();
        });
        break;

      case '4':
        const tokens = loadTokens();
        if (tokens && tokens.activeAccount) {
          console.log(`\nCurrently active account: ${tokens.activeAccount}\n`);
        } else {
          console.log('\nNo active account set.\n');
        }
        rl.close();
        showMenu();
        break;

      case '5':
        console.log('Goodbye!');
        rl.close();
        process.exit(0);
        break;

      default:
        console.log('Invalid option. Please try again.\n');
        rl.close();
        showMenu();
    }
  });
}

// Main
const args = process.argv.slice(2);
if (args.length === 0) {
  showMenu();
} else if (args[0] === '--list') {
  listAccounts();
} else if (args[0] === '--set' && args[1]) {
  switchAccount(args[1]);
} else if (args[0] === '--clear') {
  clearAllTokens();
} else {
  console.log(`
Usage: node manage-accounts.js [command]

Commands:
  --list              List all authenticated accounts
  --set <email>       Set active account
  --clear             Clear all tokens (for re-authentication)
  (no args)           Show interactive menu
  `);
}
