/**
 * Create event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');
const { validateRecurrence } = require('./recurrence');

/**
 * Create event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
const ALLOWED_SHOW_AS = ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'];

async function handleCreateEvent(args) {
  const { subject, start, end, attendees, body, isAllDay, showAs, recurrence } = args;

  if (!subject || !start || !end) {
    return {
      content: [{
        type: "text",
        text: "Subject, start, and end times are required to create an event."
      }]
    };
  }

  if (showAs !== undefined && !ALLOWED_SHOW_AS.includes(showAs)) {
    return {
      content: [{
        type: "text",
        text: `Invalid showAs value '${showAs}'. Must be one of: ${ALLOWED_SHOW_AS.join(', ')}.`
      }]
    };
  }

  const recurrenceCheck = validateRecurrence(recurrence);
  if (!recurrenceCheck.ok) {
    return { content: [{ type: "text", text: recurrenceCheck.error }] };
  }

  const startDateTime = start.dateTime || start;
  const endDateTime = end.dateTime || end;

  if (isAllDay === true) {
    const isMidnight = (dt) => typeof dt === 'string' && /T00:00(:00(\.0+)?)?$/.test(dt.replace(/[zZ]$|[+\-]\d{2}:\d{2}$/, ''));
    if (!isMidnight(startDateTime) || !isMidnight(endDateTime)) {
      return {
        content: [{
          type: "text",
          text: "All-day events require start and end to be at midnight (e.g. '2026-04-14T00:00:00') in the supplied timeZone."
        }]
      };
    }
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Build API endpoint
    const endpoint = `me/events`;

    // Request body
    const bodyContent = {
      subject,
      start: { dateTime: startDateTime, timeZone: start.timeZone || DEFAULT_TIMEZONE },
      end: { dateTime: endDateTime, timeZone: end.timeZone || DEFAULT_TIMEZONE },
      attendees: attendees?.map(email => ({ emailAddress: { address: email }, type: "required" })),
      body: { contentType: "HTML", content: body || "" }
    };

    if (isAllDay === true) bodyContent.isAllDay = true;
    if (showAs !== undefined) bodyContent.showAs = showAs;
    if (recurrence !== undefined) bodyContent.recurrence = recurrence;

    // Make API call
    const response = await callGraphAPI(accessToken, 'POST', endpoint, bodyContent);

    return {
      content: [{
        type: "text",
        text: `Event '${subject}' has been successfully created.`
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
        text: `Error creating event: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateEvent;