/**
 * What the launcher and the poller share: the api, the saved state, and the
 * cache the poller fills so the board answers without asking Graph on every
 * keystroke. Entry points `await hub.ready` first, since the state loads from
 * disk after `activate` returns.
 */

import type { ExtensionApi, ExtensionConnection, SettingValue } from "@prifly/extension-api";
import { type Access, failureText, needsSignInText, type Problem, peekAccess } from "./auth";
import { DEFAULT_COLUMNS } from "./columns";
import type { GraphFailure, Wait } from "./graph";
import type { Mail } from "./model";
import { loadState, type State, saveState } from "./state";

/** Repeats `fn` every `ms`; the returned function stops it. Tests give one they trigger by hand. */
export type Timers = { every: (fn: () => void, ms: number) => () => void };

export type HubDeps = {
  /** Sleeps out a `Retry-After`; tests give one that does not sleep. */
  wait?: Wait;
  now?: () => Date;
  /** How long the filter waits for the reader to stop typing before it searches. */
  debounceMs?: number;
  timers?: Timers;
};

export type Lists = {
  unread?: Mail[];
  inbox?: Mail[];
  flagged?: Mail[];
  /** Search columns' results, by their KQL. */
  kql: Map<string, Mail[]>;
};

/** What the last polls found in one connection's mailbox. */
export type ConnCache = {
  lists: Lists;
  /** Lower-case address of the mailbox's owner, once known. */
  address: string;
  /** Said by the last read that failed; null when it went well. */
  failure: string | null;
  /** Epoch ms of the last refresh, so a failing board is not asked again on every keystroke. */
  triedAt: number;
};

export type Hub = {
  deps: HubDeps;
  api: () => ExtensionApi | null;
  state: State;
  ready: Promise<void>;
  caches: Map<string, ConnCache>;
  /** Connections whose grant ended, learnt from a failed call while the host still lists them as fine. */
  ended: Set<string>;
  /**
   * What a refused consent said, by connection, until the scope is held:
   * its text, and the link the tenant's admin opens when the host passed one.
   */
  refused: Map<string, { text: string; url?: string }>;
  /** The newest message of a conversation when its launch was built, until `launched`. */
  launching: Map<string, string>;
  /** Search results the filter found, by connection and query. */
  searches: Map<string, Mail[]>;
  /** Counts the filter's calls so only the last of a burst searches. */
  typing: { count: number };
  /** The chips last shown, to show them again only when they change. */
  shown: { last: string };
};

export type Settings = { pollMinutes: number; columns: string; account: string };

function text(value: SettingValue | undefined, fallback: string): string {
  return typeof value === "string" ? value : fallback;
}

export function settingsOf(api: ExtensionApi): Settings {
  const minutes = api.settings?.["pollMinutes"];
  return {
    pollMinutes: typeof minutes === "number" && minutes >= 1 ? minutes : 3,
    columns: text(api.settings?.["columns"], DEFAULT_COLUMNS),
    account: text(api.settings?.["account"], ""),
  };
}

export function createHub(getApi: () => ExtensionApi | null, deps: HubDeps = {}): Hub {
  const folder = getApi()?.folder;
  const state: State = { account: "", deltas: {}, started: {} };
  const hub: Hub = {
    deps,
    api: getApi,
    state,
    ready: Promise.resolve(),
    caches: new Map(),
    ended: new Set(),
    refused: new Map(),
    launching: new Map(),
    searches: new Map(),
    typing: { count: 0 },
    shown: { last: "" },
  };
  if (folder !== undefined) {
    hub.ready = loadState(folder).then((loaded) => {
      Object.assign(state, loaded);
    });
  }
  return hub;
}

export function now(hub: Hub): Date {
  return hub.deps.now?.() ?? new Date();
}

/** Writes `state.json`; a failure is logged, not thrown. */
export async function persist(hub: Hub): Promise<void> {
  const api = hub.api();
  if (api === null) return;
  const failed = await saveState(api.folder, hub.state);
  if (failed !== null) api.log("state-not-saved", { reason: failed });
}

export function cacheFor(hub: Hub, connection: string): ConnCache {
  let cache = hub.caches.get(connection);
  if (cache === undefined) {
    cache = { lists: { kql: new Map() }, address: "", failure: null, triedAt: 0 };
    hub.caches.set(connection, cache);
  }
  return cache;
}

/**
 * The connection the board shows: the one the reader switched to, else the
 * `account` setting, else the first.
 */
export function chosenAccount(hub: Hub, api: ExtensionApi): string {
  const switched = hub.state.account;
  const exists = api.connections?.list("microsoft").some((each) => each.id === switched) ?? false;
  return exists ? switched : settingsOf(api).account;
}

/** Access to the chosen connection without ever opening the browser. */
export function boardAccess(hub: Hub, api: ExtensionApi): Access {
  const access = peekAccess(api.connections, {
    account: chosenAccount(hub, api),
    scope: "Mail.Read",
    ...(hub.deps.wait === undefined ? {} : { wait: hub.deps.wait }),
  });
  if (access.ok && hub.ended.has(access.connection.id)) {
    const problem: Problem = {
      kind: "needs-sign-in",
      text: needsSignInText(access.connection.account),
      connection: access.connection,
    };
    return { ok: false, problem };
  }
  return access;
}

/** Notes what a failed call says about the grant, for the board row and the flag. */
export function noteFailure(
  hub: Hub,
  connection: ExtensionConnection,
  failure: GraphFailure,
): void {
  if (failure.kind === "needs-sign-in" || failure.kind === "unauthorized") {
    hub.ended.add(connection.id);
  }
  if (failure.kind === "needs-sign-in" && failure.adminConsentUrl !== undefined) {
    hub.refused.set(connection.id, {
      text: failureText(failure, connection.account),
      url: failure.adminConsentUrl,
    });
  }
}

/** The reader's own addresses on a connection, lower case. */
export function selfOf(hub: Hub, connection: ExtensionConnection): string[] {
  const address = cacheFor(hub, connection.id).address;
  return [connection.account.toLowerCase(), ...(address === "" ? [] : [address])];
}
