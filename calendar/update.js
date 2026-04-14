/**
 * Update event functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');
const { validateRecurrence } = require('./recurrence');

const ALLOWED_SHOW_AS = ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'];

const isMidnight = (dt) =>
  typeof dt === 'string' && /T00:00(:00(\.0+)?)?$/.test(dt.replace(/[zZ]$|[+\-]\d{2}:\d{2}$/, ''));

async function handleUpdateEvent(args) {
  const { eventId, subject, start, end, attendees, body, location, isAllDay, showAs, recurrence } = args;

  if (!eventId) {
    return {
      content: [{ type: "text", text: "eventId is required to update an event." }]
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

  if (recurrence !== undefined && recurrence !== null) {
    const recurrenceCheck = validateRecurrence(recurrence);
    if (!recurrenceCheck.ok) {
      return { content: [{ type: "text", text: recurrenceCheck.error }] };
    }
  }

  const patch = {};
  if (subject !== undefined) patch.subject = subject;
  if (body !== undefined) patch.body = { contentType: "HTML", content: body };
  if (location !== undefined) patch.location = { displayName: location };
  if (attendees !== undefined) {
    patch.attendees = attendees.map(email => ({ emailAddress: { address: email }, type: "required" }));
  }
  if (start !== undefined) {
    patch.start = { dateTime: start.dateTime || start, timeZone: start.timeZone || DEFAULT_TIMEZONE };
  }
  if (end !== undefined) {
    patch.end = { dateTime: end.dateTime || end, timeZone: end.timeZone || DEFAULT_TIMEZONE };
  }
  if (isAllDay !== undefined) patch.isAllDay = isAllDay === true;
  if (showAs !== undefined) patch.showAs = showAs;
  if (recurrence !== undefined) patch.recurrence = recurrence; // pass null to clear the series

  if (isAllDay === true) {
    const startDt = patch.start?.dateTime;
    const endDt = patch.end?.dateTime;
    if ((startDt !== undefined && !isMidnight(startDt)) || (endDt !== undefined && !isMidnight(endDt))) {
      return {
        content: [{
          type: "text",
          text: "All-day events require start and end to be at midnight (e.g. '2026-04-14T00:00:00') in the supplied timeZone."
        }]
      };
    }
  }

  if (Object.keys(patch).length === 0) {
    return {
      content: [{ type: "text", text: "No fields to update. Provide at least one updatable property." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    const endpoint = `me/events/${eventId}`;
    await callGraphAPI(accessToken, 'PATCH', endpoint, patch);

    return {
      content: [{
        type: "text",
        text: `Event ${eventId} updated (${Object.keys(patch).join(', ')}).`
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
      content: [{ type: "text", text: `Error updating event: ${error.message}` }]
    };
  }
}

module.exports = handleUpdateEvent;
