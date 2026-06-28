/**
 * Microsoft To Do module for Outlook MCP server
 */
const handleListTodoLists = require('./lists');
const handleListTodoTasks = require('./tasks');

const todoTools = [
  {
    name: 'list-todo-lists',
    description: 'Lists Microsoft To Do task lists available to the signed-in account',
    inputSchema: {
      type: 'object',
      properties: {
        count: {
          type: 'number',
          description: 'Number of task lists to retrieve (default: 10, max: 50)'
        }
      },
      required: []
    },
    handler: handleListTodoLists,
  },
  {
    name: 'list-todo-tasks',
    description: 'Lists tasks in a Microsoft To Do task list',
    inputSchema: {
      type: 'object',
      properties: {
        listId: {
          type: 'string',
          description: 'The Microsoft To Do task list ID to inspect'
        },
        count: {
          type: 'number',
          description: 'Number of tasks to retrieve (default: 10, max: 50)'
        }
      },
      required: ['listId']
    },
    handler: handleListTodoTasks,
  }
];

module.exports = {
  todoTools,
  handleListTodoLists,
  handleListTodoTasks,
};
