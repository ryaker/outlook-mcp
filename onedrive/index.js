/**
 * OneDrive module for Outlook MCP server
 */
const handleListFiles = require('./list');
const handleSearchFiles = require('./search');
const handleDownload = require('./download');
const handleUpload = require('./upload');
const handleUploadLarge = require('./upload-large');
const handleShare = require('./share');
const { handleCreateFolder, handleDeleteItem } = require('./folder');

// OneDrive tool definitions
const onedriveTools = [
  {
    name: "onedrive-list",
    description: "List files and folders in OneDrive at a specific path",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Path to list (e.g., '/Documents', '/Photos'). Defaults to root."
        },
        count: {
          type: "number",
          description: "Number of items to retrieve (default: 25, max: 50)"
        }
      },
      required: []
    },
    handler: handleListFiles
  },
  {
    name: "onedrive-search",
    description: "Search for files in OneDrive by name or content",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query to find files"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 25, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchFiles
  },
  {
    name: "onedrive-download",
    description: "Get a download URL for a file in OneDrive. Either 'itemId' or 'path' must be provided.",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to download"
        },
        path: {
          type: "string",
          description: "Path to the file (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDownload
  },
  {
    name: "onedrive-upload",
    description: "Upload a small file (< 4MB) to OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path including filename (e.g., '/Documents/myfile.txt')"
        },
        content: {
          type: "string",
          description: "File content to upload"
        },
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path", "content"]
    },
    handler: handleUpload
  },
  {
    name: "onedrive-upload-large",
    description: "Upload a large file (> 4MB) to OneDrive using chunked upload",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Destination path including filename (e.g., '/Documents/largefile.zip')"
        },
        content: {
          type: "string",
          description: "File content to upload"
        },
        conflictBehavior: {
          type: "string",
          description: "Behavior when file exists: 'rename' (default), 'replace', or 'fail'",
          enum: ["rename", "replace", "fail"]
        }
      },
      required: ["path", "content"]
    },
    handler: handleUploadLarge
  },
  {
    name: "onedrive-share",
    description: "Create a sharing link for a file or folder in OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to share"
        },
        path: {
          type: "string",
          description: "Path to the item (alternative to itemId)"
        },
        type: {
          type: "string",
          description: "Link type: 'view' (default), 'edit', or 'embed'",
          enum: ["view", "edit", "embed"]
        },
        scope: {
          type: "string",
          description: "Link scope: 'anonymous' (default) or 'organization'",
          enum: ["anonymous", "organization"]
        }
      },
      required: []
    },
    handler: handleShare
  },
  {
    name: "onedrive-create-folder",
    description: "Create a new folder in OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        path: {
          type: "string",
          description: "Parent folder path (e.g., '/Documents'). Defaults to root."
        },
        name: {
          type: "string",
          description: "Name of the new folder"
        }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "onedrive-delete",
    description: "Delete a file or folder from OneDrive",
    inputSchema: {
      type: "object",
      properties: {
        itemId: {
          type: "string",
          description: "ID of the item to delete"
        },
        path: {
          type: "string",
          description: "Path to the item (alternative to itemId)"
        }
      },
      required: []
    },
    handler: handleDeleteItem
  }
];

module.exports = {
  onedriveTools,
  handleListFiles,
  handleSearchFiles,
  handleDownload,
  handleUpload,
  handleUploadLarge,
  handleShare,
  handleCreateFolder,
  handleDeleteItem
};
