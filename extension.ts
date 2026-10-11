/**
 * The extension as one instance: `activate` is given the api, registers the
 * MCP tools and starts the poller; the other functions answer the launcher.
 * `index.ts` exports one instance; tests make their own with a clock and
 * timers of their own.
 */

import type {
  ExtensionApi,
  ExtensionModule,
  Launch,
  LaunchBoard,
  LaunchItem,
} from "@prifly/extension-api";
import { choices } from "./choices";
import { createHub, type Hub, type HubDeps } from "./hub";
import { launch, launched } from "./launch";
import { open } from "./open";
import { type Poller, startPoller } from "./poll";
import { outlookTools } from "./tools";

export type Extension = Required<
  Pick<ExtensionModule, "activate" | "choices" | "open" | "launch" | "launched" | "action">
> & {
  /** The poller, once activated; tests run a round by hand. */
  poller: () => Poller | null;
  hub: () => Hub | null;
};

export function makeExtension(deps: HubDeps = {}): Extension {
  let api: ExtensionApi | null = null;
  let hub: Hub | null = null;
  let poller: Poller | null = null;
  const nothing: LaunchBoard = { columns: [], items: [] };
  const stopped: Launch = { prompt: "", message: "The Outlook extension is not running." };
  return {
    poller: () => poller,
    hub: () => hub,
    activate(given) {
      api = given;
      hub = createHub(() => api, deps);
      // Absent on an older prifly: optional, like every part added after the first version.
      given.tools?.register(
        outlookTools(() => api, deps.wait === undefined ? {} : { wait: deps.wait }),
      );
      poller = startPoller(hub);
      given.log("activated");
      return () => {
        poller?.stop();
        poller = null;
        hub = null;
        api = null;
      };
    },
    choices: async (launchId, query) => (hub === null ? nothing : choices(hub, launchId, query)),
    open: async (launchId, key): Promise<LaunchItem> =>
      hub === null
        ? { title: "Outlook", markdown: stopped.message ?? "" }
        : open(hub, launchId, key),
    launch: async (launchId, key, input) =>
      hub === null ? stopped : launch(hub, launchId, key, input),
    launched: async (launchId, key, sessionId) => {
      if (hub !== null) await launched(hub, launchId, key, sessionId);
    },
    // The board offers no card or chip actions in v1: nothing marks, flags or sends.
    action: () => undefined,
  };
}
