/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleUpdateEvent = require('./update');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
const { RECURRENCE_SCHEMA } = require('./recurrence');

// Calendar tool definitions
const calendarTools = [
  {
    name: "list-events",
    description: "Lists upcoming events from your calendar",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of events to retrieve (default: 10, max: 50)"
        },
        startDateTime: {
          type: "string",
          description: "ISO 8601 start date/time for the query range (default: now)"
        },
        endDateTime: {
          type: "string",
          description: "ISO 8601 end date/time for the query range (default: startDateTime + 30 days)"
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "decline-event",
    description: "Declines a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to decline"
        },
        comment: {
          type: "string",
          description: "Optional comment for declining the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "create-event",
    description: "Creates a new calendar event",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "The subject of the event"
        },
        start: {
          type: "string",
          description: "The start time of the event in ISO 8601 format"
        },
        end: {
          type: "string",
          description: "The end time of the event in ISO 8601 format"
        },
        attendees: {
          type: "array",
          items: {
            type: "string"
          },
          description: "List of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Optional body content for the event"
        },
        isAllDay: {
          type: "boolean",
          description: "Mark this event as an all-day event. Requires start/end to be at midnight in the supplied timeZone."
        },
        showAs: {
          type: "string",
          enum: ["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"],
          description: "Free/busy status to publish for this event (default: busy)"
        },
        recurrence: RECURRENCE_SCHEMA
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "update-event",
    description: "Updates fields on an existing calendar event. Only provided fields are patched.",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to update"
        },
        subject: {
          type: "string",
          description: "New subject for the event"
        },
        start: {
          type: "string",
          description: "New start time in ISO 8601 format"
        },
        end: {
          type: "string",
          description: "New end time in ISO 8601 format"
        },
        location: {
          type: "string",
          description: "New location display name"
        },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "Replacement list of attendee email addresses (replaces the existing list)"
        },
        body: {
          type: "string",
          description: "New body content (HTML)"
        },
        isAllDay: {
          type: "boolean",
          description: "Mark this event as all-day. If true, any provided start/end must be at midnight in the supplied timeZone."
        },
        showAs: {
          type: "string",
          enum: ["free", "tentative", "busy", "oof", "workingElsewhere", "unknown"],
          description: "Free/busy status to publish for this event"
        },
        recurrence: RECURRENCE_SCHEMA
      },
      required: ["eventId"]
    },
    handler: handleUpdateEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to cancel"
        },
        comment: {
          type: "string",
          description: "Optional comment for cancelling the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to delete"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleDeclineEvent,
  handleCreateEvent,
  handleUpdateEvent,
  handleCancelEvent,
  handleDeleteEvent
};
