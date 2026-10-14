/**
 * The extension's one line in Settings → Extensions: "Connected as
 * ada@example.com · 41 unread", or why it is not connected. Set once at
 * start and then at most every 15 minutes; it never opens the browser.
 */

import type { ExtensionApi } from "@prifly/extension-api";
import { failureText, peekAccess } from "./auth";
import { accountOf } from "./common";
import { unreadCount, type Wait } from "./graph";

export const STATUS_EVERY_MS = 15 * 60 * 1000;

/** Runs `fn` every `ms`; answers a function that stops it. */
export type Every = (fn: () => void, ms: number) => () => void;

const realEvery: Every = (fn, ms) => {
  const timer = setInterval(fn, ms);
  // The status line must not keep a process alive on its own.
  if (typeof timer === "object" && timer !== null && "unref" in timer) timer.unref();
  return () => clearInterval(timer);
};

/** The status line as it stands now. */
export async function statusText(api: ExtensionApi, wait?: Wait): Promise<string> {
  const access = peekAccess(api.connections, {
    account: accountOf(api, {}),
    scopes: ["Mail.Read"],
    ...(wait === undefined ? {} : { wait }),
  });
  if (!access.ok) {
    const { kind, text, connection } = access.problem;
    if (kind === "consent-needed" && connection !== null) {
      return `Connected as ${connection.account} · mail access not approved yet: the first outlook_ tool asks for it`;
    }
    return `Not connected: ${text}`;
  }
  const account = access.connection.account;
  const count = await unreadCount(access.ctx);
  if (count.ok) return `Connected as ${account} · ${count.value} unread`;
  if (count.failure.kind === "needs-sign-in" || count.failure.kind === "unauthorized") {
    return `Not connected: ${failureText(count.failure, account)}`;
  }
  return `Connected as ${account} · unread count unavailable: ${failureText(count.failure, account)}`;
}

/** Sets the status line now and every 15 minutes; answers a function that stops it. */
export function startStatus(
  api: ExtensionApi,
  options: { every?: Every; wait?: Wait } = {},
): { refresh: () => Promise<void>; stop: () => void } {
  let running: Promise<void> | null = null;
  let stopped = false;
  const refresh = (): Promise<void> => {
    if (api.status === undefined || stopped) return Promise.resolve();
    // One refresh at a time: a slow Graph never stacks them up.
    running ??= statusText(api, options.wait)
      .then((text) => {
        if (!stopped) api.status?.(text);
      })
      .catch((error: unknown) => {
        api.log("status_failed", { error: error instanceof Error ? error.message : String(error) });
      })
      .finally(() => {
        running = null;
      });
    return running;
  };
  const cancel = (options.every ?? realEvery)(() => void refresh(), STATUS_EVERY_MS);
  void refresh();
  return {
    refresh,
    stop: () => {
      stopped = true;
      cancel();
    },
  };
}
