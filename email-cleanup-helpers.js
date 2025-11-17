#!/usr/bin/env node
/**
 * Email Cleanup Helper Script
 * Provides utilities for common email cleanup tasks
 */

const https = require('https');
const path = require('path');
const fs = require('fs');

// Load TokenStorage for multi-account support
const TokenStorage = require('./auth/token-storage');
const config = require('./config');

const tokenStorage = new TokenStorage(config.AUTH_CONFIG);

// Helper to make Graph API calls
async function callGraphAPI(endpoint, method = 'GET', body = null) {
  const token = await tokenStorage.getValidAccessToken();

  if (!token) {
    throw new Error('No valid access token. Please authenticate first.');
  }

  return new Promise((resolve, reject) => {
    const url = new URL(endpoint);
    const options = {
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(url, options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const result = JSON.parse(data);
          if (res.statusCode >= 400) {
            reject(new Error(`API Error: ${res.statusCode} - ${result.error?.message || data}`));
          } else {
            resolve(result);
          }
        } catch (e) {
          reject(new Error(`Failed to parse response: ${e.message}`));
        }
      });
    });

    req.on('error', reject);
    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

// List emails with filters
async function listEmails(options = {}) {
  const {
    folder = 'inbox',
    pageSize = 25,
    filter = null,
    search = null,
    orderBy = 'receivedDateTime',
    ascending = false
  } = options;

  let endpoint = `${config.GRAPH_API_ENDPOINT}me/mailFolders/${folder}/messages?$top=${pageSize}`;

  if (search) {
    endpoint += `&$search="${encodeURIComponent(search)}"`;
  }

  if (filter) {
    endpoint += `&$filter=${encodeURIComponent(filter)}`;
  }

  endpoint += `&$orderby=${orderBy} ${ascending ? 'asc' : 'desc'}`;
  endpoint += `&$select=${config.EMAIL_SELECT_FIELDS}`;

  const result = await callGraphAPI(endpoint);
  return result.value || [];
}

// Find emails older than X days
async function findOldEmails(days = 30, folder = 'inbox') {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - days);
  const dateStr = cutoffDate.toISOString();

  const filter = `receivedDateTime lt ${dateStr}`;
  const emails = await listEmails({ folder, filter, pageSize: 50 });

  return emails;
}

// Find duplicate emails (by subject and sender)
async function findDuplicateEmails(folder = 'inbox') {
  const emails = await listEmails({ folder, pageSize: 50 });
  const seen = {};
  const duplicates = [];

  emails.forEach(email => {
    const key = `${email.from?.emailAddress?.address}|${email.subject}`;
    if (seen[key]) {
      duplicates.push({
        id: email.id,
        subject: email.subject,
        from: email.from?.emailAddress?.address,
        receivedDateTime: email.receivedDateTime,
        isDuplicate: true
      });
    } else {
      seen[key] = true;
    }
  });

  return duplicates;
}

// Get folder list
async function listFolders() {
  const result = await callGraphAPI(`${config.GRAPH_API_ENDPOINT}me/mailFolders?$top=50`);
  return result.value || [];
}

// Delete email
async function deleteEmail(emailId) {
  await callGraphAPI(
    `${config.GRAPH_API_ENDPOINT}me/messages/${emailId}`,
    'DELETE'
  );
  console.log(`✓ Deleted email: ${emailId}`);
}

// Move email to folder
async function moveEmail(emailId, folderId) {
  await callGraphAPI(
    `${config.GRAPH_API_ENDPOINT}me/messages/${emailId}/move`,
    'POST',
    { destinationId: folderId }
  );
  console.log(`✓ Moved email ${emailId} to folder ${folderId}`);
}

// Mark emails as read
async function markAsRead(emailIds) {
  for (const id of emailIds) {
    await callGraphAPI(
      `${config.GRAPH_API_ENDPOINT}me/messages/${id}`,
      'PATCH',
      { isRead: true }
    );
  }
  console.log(`✓ Marked ${emailIds.length} emails as read`);
}

// Get current active account
async function getCurrentAccount() {
  const activeAccount = await tokenStorage.getActiveAccount();
  return activeAccount;
}

// Switch active account
async function switchAccount(email) {
  await tokenStorage.setActiveAccount(email);
  console.log(`✓ Switched to account: ${email}`);
}

// Get account list
async function listAccounts() {
  const accounts = await tokenStorage.getAllAccounts();
  const active = await tokenStorage.getActiveAccount();

  console.log('\n📧 Available Accounts:');
  accounts.forEach(account => {
    const marker = account === active ? ' ✓ (ACTIVE)' : '';
    console.log(`  - ${account}${marker}`);
  });
  return accounts;
}

// CLI Interface
async function runCLI() {
  const args = process.argv.slice(2);
  const command = args[0];

  try {
    switch (command) {
      case 'list-accounts':
        await listAccounts();
        break;

      case 'switch-account':
        if (!args[1]) {
          console.error('Usage: node email-cleanup-helpers.js switch-account <email>');
          process.exit(1);
        }
        await switchAccount(args[1]);
        break;

      case 'current-account':
        const current = await getCurrentAccount();
        console.log(`Current account: ${current}`);
        break;

      case 'list-folders':
        const folders = await listFolders();
        console.log('\n📁 Available Folders:');
        folders.forEach(folder => {
          console.log(`  - ${folder.displayName} (${folder.id})`);
        });
        break;

      case 'find-old-emails':
        const days = parseInt(args[1]) || 30;
        const folder = args[2] || 'inbox';
        console.log(`\n🔍 Finding emails older than ${days} days in ${folder}...`);
        const oldEmails = await findOldEmails(days, folder);
        console.log(`\nFound ${oldEmails.length} old emails:`);
        oldEmails.forEach(email => {
          console.log(`  - ${email.subject}`);
          console.log(`    From: ${email.from?.emailAddress?.address}`);
          console.log(`    Date: ${new Date(email.receivedDateTime).toLocaleDateString()}\n`);
        });
        break;

      case 'find-duplicates':
        const dupFolder = args[1] || 'inbox';
        console.log(`\n🔍 Finding duplicate emails in ${dupFolder}...`);
        const dupes = await findDuplicateEmails(dupFolder);
        console.log(`\nFound ${dupes.length} duplicate emails:`);
        dupes.forEach(email => {
          console.log(`  - ${email.subject}`);
          console.log(`    From: ${email.from}`);
          console.log(`    ID: ${email.id}\n`);
        });
        break;

      case 'list-emails':
        const limit = parseInt(args[1]) || 25;
        console.log(`\n📧 Last ${limit} emails:`);
        const emails = await listEmails({ pageSize: limit });
        emails.forEach(email => {
          console.log(`  - ${email.subject}`);
          console.log(`    From: ${email.from?.emailAddress?.address}`);
          console.log(`    Date: ${new Date(email.receivedDateTime).toLocaleDateString()}\n`);
        });
        break;

      default:
        console.log(`
📧 Email Cleanup Helper Commands:

  list-accounts                    List all authenticated accounts
  current-account                  Show current active account
  switch-account <email>           Switch to a different account
  list-folders                     List all email folders
  list-emails [count]              List recent emails (default: 25)
  find-old-emails [days] [folder]  Find emails older than X days (default: 30 days, inbox)
  find-duplicates [folder]         Find duplicate emails in a folder

Examples:
  node email-cleanup-helpers.js list-accounts
  node email-cleanup-helpers.js switch-account pgcurtis90@hotmail.com
  node email-cleanup-helpers.js find-old-emails 90 inbox
  node email-cleanup-helpers.js find-duplicates "Sent Items"

Note: Use these commands to gather information. Then use Claude Desktop
with the Outlook MCP to perform bulk deletions or moves.
        `);
    }
  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

// Export for use as module
module.exports = {
  listEmails,
  findOldEmails,
  findDuplicateEmails,
  listFolders,
  deleteEmail,
  moveEmail,
  markAsRead,
  getCurrentAccount,
  switchAccount,
  listAccounts,
  callGraphAPI
};

// Run CLI if called directly
if (require.main === module) {
  runCLI();
}
