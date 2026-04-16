/**
 * Per-request async context — carries the target mailbox UPN through the call stack
 * so graph-api.js can rewrite `me/...` paths to `users/{upn}/...` without threading
 * the mailbox through every handler signature.
 */
const { AsyncLocalStorage } = require('async_hooks');
const config = require('../config');

const storage = new AsyncLocalStorage();

function runWithContext(ctx, fn) {
  return storage.run({ mailbox: ctx.mailbox || config.DEFAULT_MAILBOX }, fn);
}

function getMailbox() {
  const store = storage.getStore();
  return (store && store.mailbox) || config.DEFAULT_MAILBOX;
}

module.exports = { runWithContext, getMailbox };
