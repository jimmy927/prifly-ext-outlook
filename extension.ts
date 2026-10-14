/**
 * The extension as one instance: `activate` is given the api, registers the
 * MCP tools and keeps the status line. There is no board, chip or launcher:
 * everything is a tool a session calls. `index.ts` exports one instance;
 * tests make their own with timers and fetches of their own.
 */

import type { ExtensionModule } from "@prifly/extension-api";
import type { ToolDeps } from "./common";
import { type Every, startStatus } from "./status";
import { outlookTools } from "./tools";

export type ExtensionDeps = ToolDeps & {
  /** The status line's 15-minute timer; tests fire it by hand. */
  every?: Every;
};

export type Extension = Required<Pick<ExtensionModule, "activate">> & {
  /** Sets the status line again now; tests await it. */
  refreshStatus: () => Promise<void>;
};

export function makeExtension(deps: ExtensionDeps = {}): Extension {
  let refresh: () => Promise<void> = () => Promise.resolve();
  return {
    refreshStatus: () => refresh(),
    activate(given) {
      let api: typeof given | null = given;
      const { every, ...toolDeps } = deps;
      // Absent on an older prifly: optional, like every part added after the first version.
      given.tools?.register(outlookTools(() => api, toolDeps));
      const status = startStatus(given, {
        ...(every === undefined ? {} : { every }),
        ...(deps.wait === undefined ? {} : { wait: deps.wait }),
      });
      refresh = status.refresh;
      // A changed account setting shows the new account at once, not 15 minutes later.
      given.onSettings?.(() => void status.refresh());
      given.log("activated");
      return () => {
        status.stop();
        refresh = () => Promise.resolve();
        api = null;
      };
    },
  };
}
