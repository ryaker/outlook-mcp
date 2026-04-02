/**
 * Calendar module for Outlook MCP server
 */
const handleListEvents = require('./list');
const handleAcceptEvent = require('./accept');
const handleDeclineEvent = require('./decline');
const handleTentativeEvent = require('./tentative');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');

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
        },
        timezone: {
          type: "string",
          description: "IANA timezone for displayed times (e.g. 'Europe/Oslo', 'America/New_York'). Defaults to OUTLOOK_TIMEZONE env var, or UTC if unset."
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  {
    name: "accept-event",
    description: "Accepts a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to accept"
        },
        comment: {
          type: "string",
          description: "Optional comment for accepting the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleAcceptEvent
  },
  {
    name: "tentative-event",
    description: "Tentatively accepts a calendar event",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to tentatively accept"
        },
        comment: {
          type: "string",
          description: "Optional comment for tentatively accepting the event"
        }
      },
      required: ["eventId"]
    },
    handler: handleTentativeEvent
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
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
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
  handleAcceptEvent,
  handleDeclineEvent,
  handleTentativeEvent,
  handleCreateEvent,
  handleCancelEvent,
  handleDeleteEvent
};
