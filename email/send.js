/**
 * Send email functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Send email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendEmail(args) {
  const { to, cc, bcc, subject, body, importance = 'normal', saveToSentItems = true, isHtml } = args;
  
  // Validate required parameters
  if (!to) {
    return {
      content: [{ 
        type: "text", 
        text: "Recipient (to) is required."
      }]
    };
  }
  
  if (!subject) {
    return {
      content: [{ 
        type: "text", 
        text: "Subject is required."
      }]
    };
  }
  
  if (!body) {
    return {
      content: [{ 
        type: "text", 
        text: "Body content is required."
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Format recipients
    const toRecipients = to.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    });
    
    const ccRecipients = cc ? cc.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    }) : [];
    
    const bccRecipients = bcc ? bcc.split(',').map(email => {
      email = email.trim();
      return {
        emailAddress: {
          address: email
        }
      };
    }) : [];
    
    // Determine content type: explicit isHtml param takes precedence, otherwise auto-detect
    const contentType = isHtml === true ? 'html' :
                        isHtml === false ? 'text' :
                        (body.includes('<html') || body.includes('<HTML')) ? 'html' : 'text';

    // Prepare email object
    const emailObject = {
      message: {
        subject,
        body: {
          contentType: contentType,
          content: body
        },
        toRecipients,
        ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
        bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
        importance
      },
      saveToSentItems
    };
    
    // Make API call to send email
    await callGraphAPI(accessToken, 'POST', 'me/sendMail', emailObject);
    
    return {
      content: [{ 
        type: "text", 
        text: `Email sent successfully!\n\nSubject: ${subject}\nRecipients: ${toRecipients.length}${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ''}${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ''}\nMessage Length: ${body.length} characters`
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
        text: `Error sending email: ${error.message}`
      }]
    };
  }
}

module.exports = handleSendEmail;
