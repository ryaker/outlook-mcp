/**
 * Mock data functions for test mode
 */

/**
 * Simulates Microsoft Graph API responses for testing
 * @param {string} method - HTTP method
 * @param {string} path - API path
 * @param {object} data - Request data
 * @param {object} queryParams - Query parameters
 * @returns {object} - Simulated API response
 */
function simulateGraphAPIResponse(method, path, data, queryParams) {
  console.error(`Simulating response for: ${method} ${path}`);
  
  if (method === 'GET') {
    if (path.includes('/attachments/')) {
      // Single attachment download
      return {
        '@odata.type': '#microsoft.graph.fileAttachment',
        id: 'simulated-attachment-1',
        name: 'simulated-report.pdf',
        contentType: 'application/pdf',
        size: 24,
        isInline: false,
        contentBytes: Buffer.from('Simulated PDF content :)').toString('base64')
      };
    }
    if (path.includes('/attachments')) {
      // Attachment listing
      return {
        value: [
          {
            id: 'simulated-attachment-1',
            name: 'simulated-report.pdf',
            contentType: 'application/pdf',
            size: 24,
            isInline: false
          }
        ]
      };
    }
    if (path.includes('messages') && !path.includes('sendMail')) {
      // Simulate a successful email list/search response
      if (path.includes('/messages/')) {
        // Single email response
        return {
          id: "simulated-email-id",
          subject: "Simulated Email Subject",
          from: {
            emailAddress: {
              name: "Simulated Sender",
              address: "sender@example.com"
            }
          },
          toRecipients: [{
            emailAddress: {
              name: "Recipient Name",
              address: "recipient@example.com"
            }
          }],
          ccRecipients: [],
          bccRecipients: [],
          receivedDateTime: new Date().toISOString(),
          bodyPreview: "This is a simulated email preview...",
          body: {
            contentType: "text",
            content: "This is the full content of the simulated email. Since we can't connect to the real Microsoft Graph API, we're returning this placeholder content instead."
          },
          hasAttachments: true,
          importance: "normal",
          isRead: false,
          internetMessageHeaders: []
        };
      } else {
        // Email list response
        return {
          value: [
            {
              id: "simulated-email-1",
              subject: "Important Meeting Tomorrow",
              from: {
                emailAddress: {
                  name: "John Doe",
                  address: "john@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date().toISOString(),
              bodyPreview: "Let's discuss the project status...",
              hasAttachments: false,
              importance: "high",
              isRead: false
            },
            {
              id: "simulated-email-2",
              subject: "Weekly Report",
              from: {
                emailAddress: {
                  name: "Jane Smith",
                  address: "jane@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 86400000).toISOString(), // Yesterday
              bodyPreview: "Please find attached the weekly report...",
              hasAttachments: true,
              importance: "normal",
              isRead: true
            },
            {
              id: "simulated-email-3",
              subject: "Question about the project",
              from: {
                emailAddress: {
                  name: "Bob Johnson",
                  address: "bob@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 172800000).toISOString(), // 2 days ago
              bodyPreview: "I had a question about the timeline...",
              hasAttachments: false,
              importance: "normal",
              isRead: false
            }
          ]
        };
      }
    } else if (path.includes('mailFolders')) {
      // Simulate a mail folders response
      return {
        value: [
          { id: "inbox", displayName: "Inbox" },
          { id: "drafts", displayName: "Drafts" },
          { id: "sentItems", displayName: "Sent Items" },
          { id: "deleteditems", displayName: "Deleted Items" }
        ]
      };
    }
  } else if (method === 'POST' && path.includes('sendMail')) {
    // Simulate a successful email send
    return {};
  }
  
  // If we get here, we don't have a simulation for this endpoint
  console.error(`No simulation available for: ${method} ${path}`);
  return {};
}

module.exports = {
  simulateGraphAPIResponse
};
