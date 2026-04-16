/**
 * Folder management module for Outlook MCP server
 */
const handleListFolders = require('./list');
const handleCreateFolder = require('./create');
const handleMoveEmails = require('./move');

// Folder management tool definitions
const MAILBOX_PROP = {
  type: "string",
  description: "Target mailbox UPN (e.g. 'kris@cocomrepairs.com'). Defaults to DEFAULT_MAILBOX."
};

const folderTools = [
  {
    name: "list-folders",
    description: "Lists mail folders in the target mailbox.",
    inputSchema: {
      type: "object",
      properties: {
        mailbox: MAILBOX_PROP,
        includeItemCounts: { type: "boolean", description: "Include counts of total and unread items" },
        includeChildren: { type: "boolean", description: "Include child folders in hierarchy" }
      },
      required: []
    },
    handler: handleListFolders
  },
  {
    name: "create-folder",
    description: "Creates a new mail folder in the target mailbox.",
    inputSchema: {
      type: "object",
      properties: {
        mailbox: MAILBOX_PROP,
        name: { type: "string", description: "Name of the folder to create" },
        parentFolder: { type: "string", description: "Optional parent folder name (default is root)" }
      },
      required: ["name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "move-emails",
    description: "Moves emails from one folder to another within the target mailbox.",
    inputSchema: {
      type: "object",
      properties: {
        mailbox: MAILBOX_PROP,
        emailIds: { type: "string", description: "Comma-separated list of email IDs to move" },
        targetFolder: { type: "string", description: "Name of the folder to move emails to" },
        sourceFolder: { type: "string", description: "Optional name of the source folder (default is inbox)" }
      },
      required: ["emailIds", "targetFolder"]
    },
    handler: handleMoveEmails
  }
];

module.exports = {
  folderTools,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails
};
