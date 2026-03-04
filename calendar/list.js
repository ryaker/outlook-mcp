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
        // Defensive checks: handle null/undefined and string inputs
        if (!dateTimeData) return '';
        const dateTime = typeof dateTimeData === 'string' ? dateTimeData : (dateTimeData.dateTime || '');
        const timeZone = typeof dateTimeData === 'object' ? dateTimeData.timeZone : undefined;
        if (!dateTime) return '';

        // Detect if the string already contains timezone info (Z or +/-HH:MM)
        const hasOffset = /[zZ]$|[+\-]\d{2}:\d{2}$/.test(dateTime);

        // Helper to format a valid Date using the detected system timezone only if available
        const formatDateObj = (date) => {
          if (isNaN(date.getTime())) return dateTime;
          const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
          const options = {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true
          };
          if (tz) options.timeZone = tz;
          return date.toLocaleString('en-US', options);
        };

        // If it's UTC or has explicit offset, parse safely
        if (timeZone === 'UTC' || hasOffset || !timeZone) {
          const iso = dateTime.endsWith('Z') || hasOffset ? dateTime : dateTime + 'Z';
          const date = new Date(iso);
          return formatDateObj(date);
        }

        // If there's a specific timezone but the string lacks an offset, we cannot reliably convert without a timezone-aware library.
        // Return a clear fallback that includes the original timezone so we avoid silently showing an incorrect local time.
        return `${dateTime} (${timeZone})`;
      };

      const startDate = formatDateTime(event.start);
      const endDate = formatDateTime(event.end);
      const location = event.location?.displayName || 'No location';
      
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
