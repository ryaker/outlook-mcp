/**
 * List Microsoft To Do task lists functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

async function handleListTodoLists(args = {}) {
  const count = Math.min(args.count || 10, config.MAX_RESULT_COUNT);

  try {
    const accessToken = await ensureAuthenticated();
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/todo/lists',
      null,
      {
        $top: count,
      }
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: 'text', text: 'No Microsoft To Do task lists found.' }]
      };
    }

    const listText = response.value.map((list, index) => {
      const wellKnown = list.wellknownListName && list.wellknownListName !== 'none'
        ? `\nWell-known list: ${list.wellknownListName}`
        : '';
      const ownerStatus = typeof list.isOwner === 'boolean'
        ? `\nOwner: ${list.isOwner ? 'yes' : 'no'}`
        : '';

      return `${index + 1}. ${list.displayName || 'Untitled list'}${ownerStatus}${wellKnown}\nID: ${list.id}\n`;
    }).join('\n');

    return {
      content: [{
        type: 'text',
        text: `Found ${response.value.length} Microsoft To Do task lists:\n\n${listText}`
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
      content: [{ type: 'text', text: `Error listing Microsoft To Do task lists: ${error.message}` }]
    };
  }
}

module.exports = handleListTodoLists;
