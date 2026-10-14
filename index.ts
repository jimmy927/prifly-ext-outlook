/**
 * The Outlook extension: what prifly imports. `activate` is given the api,
 * registers the MCP tools and keeps the status line. Types come from
 * `@prifly/extension-api` with `import type`, so nothing of it is loaded at
 * run time.
 */

import { makeExtension } from "./extension";

export const activate = makeExtension().activate;
