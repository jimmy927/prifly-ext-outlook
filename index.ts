/**
 * The Outlook extension: what prifly imports. `activate` is given the api,
 * registers the MCP tools and starts the poller; the other exports answer the
 * "From an Outlook mail…" launcher. Types come from `@prifly/extension-api`
 * with `import type`, so nothing of it is loaded at run time.
 */

import { makeExtension } from "./extension";

const extension = makeExtension();

export const activate = extension.activate;
export const choices = extension.choices;
export const open = extension.open;
export const launch = extension.launch;
export const launched = extension.launched;
export const action = extension.action;
