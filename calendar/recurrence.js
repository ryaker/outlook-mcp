/**
 * Recurrence validation for calendar events.
 * Mirrors Microsoft Graph's patternedRecurrence shape:
 *   { pattern: { type, interval, ... }, range: { type, startDate, ... } }
 */

const PATTERN_TYPES = [
  'daily',
  'weekly',
  'absoluteMonthly',
  'relativeMonthly',
  'absoluteYearly',
  'relativeYearly'
];

const RANGE_TYPES = ['endDate', 'noEnd', 'numbered'];

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

function validateRecurrence(recurrence) {
  if (recurrence === undefined || recurrence === null) return { ok: true };
  if (typeof recurrence !== 'object') {
    return { ok: false, error: "recurrence must be an object with 'pattern' and 'range'." };
  }

  const { pattern, range } = recurrence;
  if (!pattern || typeof pattern !== 'object') {
    return { ok: false, error: "recurrence.pattern is required." };
  }
  if (!PATTERN_TYPES.includes(pattern.type)) {
    return { ok: false, error: `recurrence.pattern.type must be one of: ${PATTERN_TYPES.join(', ')}.` };
  }
  if (typeof pattern.interval !== 'number' || pattern.interval < 1) {
    return { ok: false, error: "recurrence.pattern.interval must be a positive integer." };
  }

  if (!range || typeof range !== 'object') {
    return { ok: false, error: "recurrence.range is required." };
  }
  if (!RANGE_TYPES.includes(range.type)) {
    return { ok: false, error: `recurrence.range.type must be one of: ${RANGE_TYPES.join(', ')}.` };
  }
  if (!ISO_DATE.test(range.startDate || '')) {
    return { ok: false, error: "recurrence.range.startDate must be 'YYYY-MM-DD'." };
  }
  if (range.type === 'endDate' && !ISO_DATE.test(range.endDate || '')) {
    return { ok: false, error: "recurrence.range.endDate must be 'YYYY-MM-DD' when range.type is 'endDate'." };
  }
  if (range.type === 'numbered' && (typeof range.numberOfOccurrences !== 'number' || range.numberOfOccurrences < 1)) {
    return { ok: false, error: "recurrence.range.numberOfOccurrences must be a positive integer when range.type is 'numbered'." };
  }

  return { ok: true };
}

const RECURRENCE_SCHEMA = {
  type: "object",
  description: "Graph patternedRecurrence. Example: {\"pattern\":{\"type\":\"weekly\",\"interval\":1,\"daysOfWeek\":[\"monday\"]},\"range\":{\"type\":\"endDate\",\"startDate\":\"2026-04-14\",\"endDate\":\"2026-07-14\"}}",
  properties: {
    pattern: {
      type: "object",
      properties: {
        type: { type: "string", enum: PATTERN_TYPES },
        interval: { type: "number" },
        daysOfWeek: { type: "array", items: { type: "string" } },
        dayOfMonth: { type: "number" },
        month: { type: "number" },
        firstDayOfWeek: { type: "string" },
        index: { type: "string" }
      },
      required: ["type", "interval"]
    },
    range: {
      type: "object",
      properties: {
        type: { type: "string", enum: RANGE_TYPES },
        startDate: { type: "string", description: "YYYY-MM-DD" },
        endDate: { type: "string", description: "YYYY-MM-DD (required when range.type is 'endDate')" },
        numberOfOccurrences: { type: "number" },
        recurrenceTimeZone: { type: "string" }
      },
      required: ["type", "startDate"]
    }
  },
  required: ["pattern", "range"]
};

module.exports = { validateRecurrence, RECURRENCE_SCHEMA };
