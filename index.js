#!/usr/bin/env node
/**
 * CCR Outlook MCP Server — HTTP (StreamableHTTP) transport for Railway.
 * App-only Microsoft Graph auth via MSAL client credentials.
 */
const express = require('express');
const cors = require('cors');
const { randomUUID } = require('crypto');
const { Server } = require('@modelcontextprotocol/sdk/server/index.js');
const { StreamableHTTPServerTransport } = require('@modelcontextprotocol/sdk/server/streamableHttp.js');

const config = require('./config');
const { runWithContext } = require('./utils/context');

const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');

const TOOLS = [...emailTools, ...folderTools, ...rulesTools];

console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} v${config.SERVER_VERSION}`);
console.error(`Default mailbox: ${config.DEFAULT_MAILBOX}`);
console.error(`Tools registered: ${TOOLS.map((t) => t.name).join(', ')}`);

function buildServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    const { method, params, id } = request;
    console.error(`REQUEST: ${method} [${id}]`);

    if (method === 'initialize') {
      return {
        protocolVersion: '2025-06-18',
        capabilities: { tools: {} },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION },
      };
    }

    if (method === 'tools/list') {
      return {
        tools: TOOLS.map((t) => ({
          name: t.name,
          description: t.description,
          inputSchema: t.inputSchema,
        })),
      };
    }

    if (method === 'resources/list') return { resources: [] };
    if (method === 'prompts/list') return { prompts: [] };

    if (method === 'tools/call') {
      try {
        const { name, arguments: args = {} } = params || {};
        const tool = TOOLS.find((t) => t.name === name);
        if (!tool || !tool.handler) {
          return { error: { code: -32601, message: `Tool not found: ${name}` } };
        }
        // Run the handler inside the per-request mailbox context so graph-api.js
        // rewrites `me/...` → `users/{mailbox}/...` correctly.
        return await runWithContext({ mailbox: args.mailbox }, () => tool.handler(args));
      } catch (error) {
        console.error('Error in tools/call:', error);
        return { error: { code: -32603, message: `Error processing tool call: ${error.message}` } };
      }
    }

    return { error: { code: -32601, message: `Method not found: ${method}` } };
  };

  return server;
}

const app = express();
app.use(cors({ origin: '*', exposedHeaders: ['mcp-session-id'] }));
app.use(express.json({ limit: '4mb' }));

app.get('/health', (_req, res) => {
  res.json({
    ok: true,
    server: config.SERVER_NAME,
    version: config.SERVER_VERSION,
    defaultMailbox: config.DEFAULT_MAILBOX,
    tools: TOOLS.length,
  });
});

app.get('/', (_req, res) => {
  res.type('text/plain').send(
    `${config.SERVER_NAME} v${config.SERVER_VERSION}\n` +
      `MCP endpoint: POST /mcp\n` +
      `Health: GET /health\n` +
      `Default mailbox: ${config.DEFAULT_MAILBOX}\n`
  );
});

// Stateless mode: every POST spins up a fresh Server+Transport pair. Simpler and
// resilient across Railway autoscale/redeploys than session-ID stickiness.
app.post('/mcp', async (req, res) => {
  try {
    const server = buildServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
    });
    res.on('close', () => {
      transport.close();
      server.close();
    });
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err) {
    console.error('MCP request error:', err);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: { code: -32603, message: 'Internal server error' },
        id: null,
      });
    }
  }
});

app.get('/mcp', (_req, res) => {
  res.status(405).json({
    jsonrpc: '2.0',
    error: { code: -32000, message: 'Method Not Allowed. Use POST for MCP requests.' },
    id: null,
  });
});

app.delete('/mcp', (_req, res) => {
  res.status(405).json({
    jsonrpc: '2.0',
    error: { code: -32000, message: 'Method Not Allowed.' },
    id: null,
  });
});

const PORT = config.PORT;
app.listen(PORT, '0.0.0.0', () => {
  console.error(`${config.SERVER_NAME} listening on 0.0.0.0:${PORT}`);
});

process.on('SIGTERM', () => {
  console.error('SIGTERM received; exiting.');
  process.exit(0);
});
