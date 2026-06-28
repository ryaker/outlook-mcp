/**
 * List Microsoft To Do tasks functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

function formatGraphDateTime(dateTimeValue) {
  if (!dateTimeValue || !dateTimeValue.dateTime) {
    return 'None';
  }

  const raw = dateTimeValue.dateTime;
  const zone = dateTimeValue.timeZone;
  const hasOffset = /[zZ]$|[+\-]\d{2}:\d{2}$/.test(raw);
  const iso = hasOffset ? raw : `${raw}Z`;
  const parsed = new Date(iso);

  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toLocaleString();
  }

  return zone ? `${raw} (${zone})` : raw;
}

async function handleListTodoTasks(args = {}) {
  const listId = args.listId;
  if (!listId) {
    return {
      content: [{ type: 'text', text: 'A Microsoft To Do listId is required.' }]
    };
  }

  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);

  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/todo/lists/${listId}/tasks`,
      null,
      {
        $top: count,
        $select: config.TODO_TASK_SELECT_FIELDS,
      }
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: 'text', text: `No Microsoft To Do tasks found in list ${listId}.` }]
      };
    }

    const taskText = response.value.map((task, index) => {
      const due = formatGraphDateTime(task.dueDateTime);
      const completed = formatGraphDateTime(task.completedDateTime);
      const importance = task.importance || 'normal';
      const body = task.body?.content ? `\nNotes: ${task.body.content}` : '';

      return `${index + 1}. ${task.title || 'Untitled task'}\nStatus: ${task.status || 'unknown'}\nImportance: ${importance}\nDue: ${due}\nCompleted: ${completed}${body}\nID: ${task.id}\n`;
    }).join('\n');

    return {
      content: [{
        type: 'text',
        text: `Found ${response.value.length} tasks in Microsoft To Do list ${listId}:\n\n${taskText}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: 'text', text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }

    if (error.message === 'UNAUTHORIZED') {
      return {
        content: [{
          type: 'text',
          text: 'Microsoft To Do access is not authorized for the current token. Re-authenticate after granting Tasks.ReadWrite.'
        }]
      };
    }

    return {
      content: [{ type: 'text', text: `Error listing Microsoft To Do tasks: ${error.message}` }]
    };
  }
}

module.exports = handleListTodoTasks;
