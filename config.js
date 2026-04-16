/**
 * Configuration for CCR Outlook MCP Server (app-only / client credentials flow)
 */
module.exports = {
  SERVER_NAME: "ccr-outlook-mcp",
  SERVER_VERSION: "3.0.0",

  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',

  AUTH_CONFIG: {
    clientId: process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.MS_CLIENT_SECRET || '',
    tenantId: process.env.MS_TENANT_ID || '',
    scopes: ['https://graph.microsoft.com/.default'],
  },

  DEFAULT_MAILBOX: process.env.DEFAULT_MAILBOX || 'kris@cocomrepairs.com',

  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',

  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',

  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  PORT: parseInt(process.env.PORT, 10) || 8080,
};
