/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Build API endpoint
    let endpoint = 'me/events';
    
    // Add query parameters
    const queryParams = {
      $top: count,
      $orderby: 'start/dateTime',
      $filter: `start/dateTime ge '${new Date().toISOString()}'`,
      $select: config.CALENDAR_SELECT_FIELDS
    };
    
    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No calendar events found."
        }]
      };
    }
    
    // Detect system timezone
    const systemTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    
    // Format results
    const eventList = response.value.map((event, index) => {
      const formatDateTime = (dateTimeData) => {
        try {
          const { dateTime, timeZone } = dateTimeData;
          
          // Parse the date. If timeZone is UTC, ensure JS treats it as UTC.
          let date;
          if (timeZone === 'UTC' || !timeZone) {
            // Append Z if missing to force UTC parsing
            date = new Date(dateTime.endsWith('Z') ? dateTime : dateTime + 'Z');
          } else {
            // If it's a specific timezone (though usually UTC without Prefer header)
            // Note: JS doesn't easily parse non-UTC strings without a library like luxon
            // but we'll try our best.
            date = new Date(dateTime);
          }

          if (isNaN(date.getTime())) return dateTime;

          // Format for display using the system timezone (user's local time)
          return date.toLocaleString(undefined, {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true,
            timeZone: systemTimezone
          });
        } catch (e) {
          return dateTimeData.dateTime;
        }
      };

      const startDate = formatDateTime(event.start);
      const endDate = formatDateTime(event.end);
      const location = event.location.displayName || 'No location';
      
      return `${index + 1}. ${event.subject} - Location: ${location}\nStart: ${startDate}\nEnd: ${endDate}\nSummary: ${event.bodyPreview}\nID: ${event.id}\n`;
    }).join("\n");

    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} events:\n\n${eventList}`
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
        text: `Error listing events: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEvents;
