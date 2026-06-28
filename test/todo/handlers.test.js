const handleListTodoLists = require('../../todo/lists');
const handleListTodoTasks = require('../../todo/tasks');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('Microsoft To Do handlers', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
  });

  describe('handleListTodoLists', () => {
    test('lists available task lists', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [
          {
            id: 'list-1',
            displayName: 'Tasks',
            isOwner: true,
            wellknownListName: 'defaultList'
          },
          {
            id: 'list-2',
            displayName: 'Work',
            isOwner: true,
            wellknownListName: 'none'
          }
        ]
      });

      const result = await handleListTodoLists({ count: 5 });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/todo/lists',
        null,
        expect.objectContaining({ $top: 5 })
      );
      expect(result.content[0].text).toContain('Found 2 Microsoft To Do task lists');
      expect(result.content[0].text).toContain('Tasks');
      expect(result.content[0].text).toContain('ID: list-1');
    });

    test('returns a helpful empty-state message when no task lists exist', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListTodoLists({});

      expect(result.content[0].text).toBe('No Microsoft To Do task lists found.');
    });

    test('returns an authentication prompt when auth is missing', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleListTodoLists({});

      expect(result.content[0].text).toBe("Authentication required. Please use the 'authenticate' tool first.");
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('surfaces missing To Do permissions clearly when Graph rejects the request', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockRejectedValue(new Error('UNAUTHORIZED'));

      const result = await handleListTodoLists({});

      expect(result.content[0].text).toContain('Microsoft To Do access is not authorized for the current token');
    });
  });

  describe('handleListTodoTasks', () => {
    test('lists tasks from a task list by id', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({
        value: [
          {
            id: 'task-1',
            title: 'Pay rent',
            status: 'notStarted',
            importance: 'high',
            dueDateTime: { dateTime: '2026-07-01T21:00:00.0000000', timeZone: 'UTC' },
            body: { content: 'Before the first of the month' }
          },
          {
            id: 'task-2',
            title: 'Buy groceries',
            status: 'completed',
            importance: 'normal',
            completedDateTime: { dateTime: '2026-06-27T18:00:00.0000000', timeZone: 'UTC' },
            body: { content: '' }
          }
        ]
      });

      const result = await handleListTodoTasks({ listId: 'list-1', count: 10 });

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken,
        'GET',
        'me/todo/lists/list-1/tasks',
        null,
        expect.objectContaining({ $top: 10 })
      );
      expect(result.content[0].text).toContain('Found 2 tasks in Microsoft To Do list list-1');
      expect(result.content[0].text).toContain('Pay rent');
      expect(result.content[0].text).toContain('Status: completed');
      expect(result.content[0].text).toContain('ID: task-2');
    });

    test('requires a task list id', async () => {
      const result = await handleListTodoTasks({});

      expect(result.content[0].text).toBe('A Microsoft To Do listId is required.');
      expect(ensureAuthenticated).not.toHaveBeenCalled();
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    test('returns a helpful empty-state message when a task list has no tasks', async () => {
      ensureAuthenticated.mockResolvedValue(mockAccessToken);
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await handleListTodoTasks({ listId: 'list-empty' });

      expect(result.content[0].text).toBe('No Microsoft To Do tasks found in list list-empty.');
    });

    test('returns an authentication prompt when auth is missing', async () => {
      ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

      const result = await handleListTodoTasks({ listId: 'list-1' });

      expect(result.content[0].text).toBe("Authentication required. Please use the 'authenticate' tool first.");
      expect(callGraphAPI).not.toHaveBeenCalled();
    });
  });
});
