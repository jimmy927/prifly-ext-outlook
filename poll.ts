/**
 * The poller: every `pollMinutes`, for each Microsoft connection that holds
 * Mail.Read, it reads what changed in the inbox (a delta from the saved
 * link), refreshes the board's cache, and from that lights the launcher and
 * puts chips on sessions whose conversation got a new message. It never
 * opens the browser and never throws.
 */

import type { ExtensionApi, ExtensionConnection } from "@prifly/extension-api";
import { failureText, peekAccess } from "./auth";
import { parseColumns } from "./columns";
import { type GraphContext, type InboxDelta, inboxDelta, me } from "./graph";
import { cacheFor, type Hub, noteFailure, now, persist, settingsOf } from "./hub";
import { type Connected, isCold, type Mode, refreshLists } from "./refresh";
import { noteReplies, renderChips, updateFlag } from "./signals";

/** How far back the first delta reaches; older mail is not on the board. */
const SEED_DAYS = 30;
const DAY_MS = 86_400_000;

export type Poller = {
  /** One round over every connection; resolves when it is done. */
  tick: () => Promise<void>;
  /** Stops the timer. */
  stop: () => void;
};

type Changes = InboxDelta & { seeded: boolean };

/** The inbox's changes since the saved link; a link that expired starts over. Null when Graph failed. */
async function readDelta(hub: Hub, on: Connected): Promise<Changes | null> {
  const id = on.connection.id;
  const since = new Date(now(hub).getTime() - SEED_DAYS * DAY_MS).toISOString();
  const link = hub.state.deltas[id];
  let seeded = link === undefined;
  let result = await inboxDelta(on.ctx, link, link === undefined ? since : undefined);
  if (!result.ok && result.failure.kind === "gone") {
    // The link expired (410 or syncStateNotFound): forget it and read the recent inbox afresh.
    delete hub.state.deltas[id];
    seeded = true;
    result = await inboxDelta(on.ctx, undefined, since);
  }
  if (!result.ok) {
    noteFailure(hub, on.connection, result.failure);
    cacheFor(hub, id).failure = failureText(result.failure, on.connection.account);
    return null;
  }
  hub.state.deltas[id] = result.value.deltaLink;
  return { ...result.value, seeded };
}

/** The mailbox owner's address, so the reader's own mail is not taken for a reply. */
async function learnAddress(hub: Hub, ctx: GraphContext, connection: string): Promise<void> {
  const cache = cacheFor(hub, connection);
  if (cache.address !== "") return;
  const profile = await me(ctx);
  if (profile.ok) cache.address = profile.value.address.toLowerCase();
}

async function pollConnection(
  hub: Hub,
  api: ExtensionApi,
  connection: ExtensionConnection,
): Promise<void> {
  const access = peekAccess(api.connections, {
    account: connection.id,
    scope: "Mail.Read",
    ...(hub.deps.wait === undefined ? {} : { wait: hub.deps.wait }),
  });
  // No Mail.Read or a grant that ended: the board says so in a row; nothing to poll.
  if (!access.ok) return;
  const on: Connected = { ctx: access.ctx, connection: access.connection };
  await learnAddress(hub, on.ctx, connection.id);
  const changes = await readDelta(hub, on);
  if (changes === null) return;
  const cache = cacheFor(hub, connection.id);
  const self = [connection.account.toLowerCase(), cache.address];
  noteReplies(hub, connection.id, changes.changed, self);
  const columns = parseColumns(settingsOf(api).columns);
  const moved = changes.seeded || changes.changed.length > 0 || changes.removed.length > 0;
  const mode: Mode = moved || isCold(hub, connection.id, columns) ? "all" : "flagged";
  await refreshLists(hub, on, columns, mode);
}

/**
 * Starts polling: a first round now, then one every `pollMinutes` (read again
 * when the reader changes the setting). `timers` is injectable for tests.
 */
export function startPoller(hub: Hub, options: { immediate?: boolean } = {}): Poller {
  const round = async () => {
    const api = hub.api();
    if (api === null || api.connections === undefined) return;
    try {
      await hub.ready;
      for (const connection of api.connections.list("microsoft")) {
        await pollConnection(hub, api, connection);
      }
      updateFlag(hub);
      renderChips(hub);
      await persist(hub);
    } catch (error) {
      api.log("poll-failed", { reason: error instanceof Error ? error.message : String(error) });
    }
  };
  // A round already under way is the answer to asking for another.
  let current: Promise<void> | null = null;
  const tick = () => {
    current ??= round().finally(() => {
      current = null;
    });
    return current;
  };
  const every = hub.deps.timers?.every ?? defaultEvery;
  let stopTimer = () => {};
  let stopped = false;
  const schedule = () => {
    stopTimer();
    const api = hub.api();
    if (stopped) return;
    const minutes = api === null ? 3 : settingsOf(api).pollMinutes;
    stopTimer = every(() => void tick(), minutes * 60_000);
  };
  schedule();
  hub.api()?.onSettings?.(schedule);
  if (options.immediate !== false) void tick();
  return {
    tick,
    stop: () => {
      stopped = true;
      stopTimer();
      stopTimer = () => {};
    },
  };
}

function defaultEvery(fn: () => void, ms: number): () => void {
  const handle = setInterval(fn, ms);
  // A timer alone must not keep the host's process alive.
  handle.unref();
  return () => clearInterval(handle);
}
