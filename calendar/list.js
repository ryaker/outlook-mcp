/**
 * List events functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Format a Graph API dateTime string for display.
 * When Prefer: outlook.timezone header is set, dateTime is in that timezone.
 * Input: "2026-03-16T18:30:00.0000000" → Output: "3/16/2026, 6:30:00 PM"
 */
function formatGraphDateTime(dateTimeStr) {
  if (!dateTimeStr || !dateTimeStr.includes('T')) return dateTimeStr || 'N/A';

  const [datePart, timePart] = dateTimeStr.split('T');
  if (!datePart || !timePart) return dateTimeStr;

  const [year, month, day] = datePart.split('-').map(Number);
  const [hours, minutes] = timePart.split(':').map(Number);
  if (isNaN(year) || isNaN(hours)) return dateTimeStr;

  const h12 = hours % 12 || 12;
  const ampm = hours < 12 ? 'AM' : 'PM';
  const mm = String(minutes).padStart(2, '0');

  return `${month}/${day}/${year}, ${h12}:${mm}:00 ${ampm}`;
}

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);
  const timezone = args.timezone || config.DISPLAY_TIMEZONE;

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Use calendarView endpoint to include expanded recurring event instances
    const startDate = args.startDateTime ? new Date(args.startDateTime) : new Date();
    const endDate = args.endDateTime
      ? new Date(args.endDateTime)
      : new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000); // default 30 days

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || endDate <= startDate) {
      return {
        content: [{
          type: "text",
          text: "Invalid date range. Provide valid startDateTime/endDateTime with endDateTime > startDateTime."
        }]
      };
    }

    let endpoint = 'me/calendarView';

    // calendarView requires startDateTime and endDateTime as query params
    const queryParams = {
      startDateTime: startDate.toISOString(),
      endDateTime: endDate.toISOString(),
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS
    };

    // Request times in the specified timezone so dateTime values are local
    const extraHeaders = timezone
      ? { 'Prefer': `outlook.timezone="${timezone}"` }
      : {};

    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams, extraHeaders);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No calendar events found."
        }]
      };
    }

    // Format results
    // Note: With Prefer: outlook.timezone header, Graph API returns dateTime in the requested timezone.
    // Do NOT pass it through new Date() which would treat it as UTC and shift it.
    const eventList = response.value.map((event, index) => {
      const startDate = formatGraphDateTime(event.start?.dateTime || '');
      const endDate = formatGraphDateTime(event.end?.dateTime || '');
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
