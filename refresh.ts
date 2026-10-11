/**
 * Fills a connection's cache from Graph: the unread, inbox and flagged lists
 * and each search column's hits. The poller does it on a timer; the board
 * does it once when it opens on an empty cache. Never throws: a failed read
 * is text in `cache.failure`, and the lists already held stay.
 */

import type { ExtensionConnection } from "@prifly/extension-api";
import { failureText } from "./auth";
import type { ColumnSpec } from "./columns";
import { flagged, type GraphContext, type GraphResult, inbox, search, unread } from "./graph";
import { cacheFor, type Hub, noteFailure, now } from "./hub";
import type { Mail } from "./model";

const TOP_UNREAD = 100;
const TOP_INBOX = 50;
const TOP_FLAGGED = 50;
const TOP_SEARCH = 30;

export type Connected = { ctx: GraphContext; connection: ExtensionConnection };

/**
 * What to fetch: `missing` only what the cache lacks (the board opened cold,
 * or a column was added), `all` everything (a poll that saw the mailbox
 * change), `flagged` the flagged list alone (a quiet poll; flagged mail in
 * other folders is not in the inbox's delta).
 */
export type Mode = "missing" | "all" | "flagged";

type ListName = "unread" | "inbox" | "flagged" | "kql";

/** The KQL searches the columns ask for. */
export function searchesOf(columns: ColumnSpec[]): string[] {
  return columns.flatMap((column) => (column.source.kind === "kql" ? [column.source.kql] : []));
}

/** Whether any list a board needs is not in the cache yet. */
export function isCold(hub: Hub, connection: string, columns: ColumnSpec[]): boolean {
  const { lists } = cacheFor(hub, connection);
  const missing =
    lists.unread === undefined || lists.inbox === undefined || lists.flagged === undefined;
  return missing || searchesOf(columns).some((kql) => !lists.kql.has(kql));
}

/** Refreshes the lists `mode` names. */
export async function refreshLists(
  hub: Hub,
  on: Connected,
  columns: ColumnSpec[],
  mode: Mode,
): Promise<void> {
  const cache = cacheFor(hub, on.connection.id);
  const { lists } = cache;
  const failures: string[] = [];
  const wants = (list: ListName, held: boolean) =>
    mode === "all" || (mode === "missing" && !held) || (mode === "flagged" && list === "flagged");
  const take = async (
    wanted: boolean,
    read: () => Promise<GraphResult<Mail[]>>,
    keep: (mails: Mail[]) => void,
  ) => {
    if (!wanted) return;
    const result = await read();
    if (result.ok) {
      keep(result.value);
      return;
    }
    noteFailure(hub, on.connection, result.failure);
    failures.push(failureText(result.failure, on.connection.account));
  };
  await take(
    wants("unread", lists.unread !== undefined),
    () => unread(on.ctx, TOP_UNREAD),
    (mails) => {
      lists.unread = mails;
    },
  );
  await take(
    wants("inbox", lists.inbox !== undefined),
    () => inbox(on.ctx, TOP_INBOX),
    (mails) => {
      lists.inbox = mails;
    },
  );
  await take(
    wants("flagged", lists.flagged !== undefined),
    () => flagged(on.ctx, TOP_FLAGGED),
    (mails) => {
      lists.flagged = mails;
    },
  );
  for (const kql of searchesOf(columns)) {
    await take(
      wants("kql", lists.kql.has(kql)),
      () => search(on.ctx, kql, TOP_SEARCH),
      (mails) => {
        lists.kql.set(kql, mails);
      },
    );
  }
  cache.failure = failures[0] ?? null;
  cache.triedAt = now(hub).getTime();
  if (failures.length === 0) hub.ended.delete(on.connection.id);
}
