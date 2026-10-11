/**
 * What the launcher lists. With nothing typed, the board comes from the
 * cache the poller fills (Graph is asked only when the cache is empty). With
 * something typed, the cards that match show at once, and once the reader
 * has stopped typing for `debounceMs` a KQL search answers instead.
 */

import type { ExtensionApi, ExtensionConnection, LaunchBoard } from "@prifly/extension-api";
import { failureText } from "./auth";
import { buildBoard, cardsOf, narrowed } from "./board";
import { type ColumnSpec, parseColumns } from "./columns";
import { knownConversations } from "./conversations";
import { search } from "./graph";
import { boardAccess, cacheFor, type Hub, now, selfOf, settingsOf } from "./hub";
import type { Mail } from "./model";
import { type Connected, isCold, refreshLists } from "./refresh";
import { accessBoard, boardActions, failureBoard } from "./rows";
import { LAUNCHER_ID, updateFlag } from "./signals";

/** A board that failed to load is asked again no sooner than this. */
const RETRY_MS = 30_000;
const TOP_SEARCH = 30;
const KEEP_SEARCHES = 20;
const DEFAULT_DEBOUNCE_MS = 300;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function listFor(hub: Hub, connection: string, column: ColumnSpec): Mail[] {
  const { lists } = cacheFor(hub, connection);
  const { source } = column;
  if (source.kind === "kql") return lists.kql.get(source.kql) ?? [];
  return lists[source.preset] ?? [];
}

function allLists(hub: Hub, connection: string): Mail[][] {
  const { lists } = cacheFor(hub, connection);
  return [lists.unread ?? [], lists.inbox ?? [], lists.flagged ?? [], ...lists.kql.values()];
}

function boardOf(
  hub: Hub,
  api: ExtensionApi,
  connection: ExtensionConnection,
  columns: ColumnSpec[],
): LaunchBoard {
  return buildBoard({
    columns: columns.map((column) => ({
      id: column.id,
      name: column.name,
      mails: listFor(hub, connection.id, column),
    })),
    lists: allLists(hub, connection.id),
    started: new Set(Object.keys(hub.state.started)),
    self: selfOf(hub, connection),
    now: now(hub),
    actions: boardActions(api),
  });
}

/** The hits of a search column, one column of cards. */
function searchBoard(
  hub: Hub,
  api: ExtensionApi,
  connection: ExtensionConnection,
  query: string,
  hits: Mail[],
): LaunchBoard {
  const input = {
    started: new Set(Object.keys(hub.state.started)),
    self: selfOf(hub, connection),
    now: now(hub),
    actions: boardActions(api),
  };
  const known = knownConversations([...allLists(hub, connection.id), hits]);
  return {
    columns: [{ id: "search", name: `Search: ${query}` }],
    items: cardsOf(hits, "search", input, known),
    actions: input.actions,
  };
}

/** The search for what was typed, asked once and remembered. */
async function hitsFor(hub: Hub, on: Connected, query: string) {
  const key = `${on.connection.id}|${query}`;
  const held = hub.searches.get(key);
  if (held !== undefined) return { ok: true as const, value: held };
  const found = await search(on.ctx, query, TOP_SEARCH);
  if (!found.ok) return found;
  if (hub.searches.size >= KEEP_SEARCHES) {
    const oldest = hub.searches.keys().next().value;
    if (oldest !== undefined) hub.searches.delete(oldest);
  }
  hub.searches.set(key, found.value);
  return found;
}

/** The board narrowed by what was typed from hand now, then by a search once typing stops. */
async function searched(
  hub: Hub,
  api: ExtensionApi,
  on: Connected,
  columns: ColumnSpec[],
  query: string,
  mine: number,
): Promise<LaunchBoard> {
  await sleep(hub.deps.debounceMs ?? DEFAULT_DEBOUNCE_MS);
  // The reader typed on: this answer is stale, so give what is in hand and ask nothing.
  if (hub.typing.count !== mine) return narrowed(boardOf(hub, api, on.connection, columns), query);
  const found = await hitsFor(hub, on, query);
  if (!found.ok) return failureBoard(api, failureText(found.failure, on.connection.account));
  return searchBoard(hub, api, on.connection, query, found.value);
}

/** Whether the cache wants Graph: it is empty and was not tried a moment ago. */
function wantsRefresh(hub: Hub, connection: string, columns: ColumnSpec[]): boolean {
  if (!isCold(hub, connection, columns)) return false;
  const { failure, triedAt } = cacheFor(hub, connection);
  return failure === null || now(hub).getTime() - triedAt > RETRY_MS;
}

export async function choices(hub: Hub, launchId: string, query: string): Promise<LaunchBoard> {
  // Counted before anything is awaited, so a burst of keystrokes is told apart.
  const mine = ++hub.typing.count;
  await hub.ready;
  const api = hub.api();
  if (api === null || launchId !== LAUNCHER_ID) return { columns: [], items: [] };
  const access = boardAccess(hub, api);
  if (!access.ok) return accessBoard(hub, api, access.problem);
  const on: Connected = { ctx: access.ctx, connection: access.connection };
  const columns = parseColumns(settingsOf(api).columns);
  if (wantsRefresh(hub, on.connection.id, columns)) {
    await refreshLists(hub, on, columns, "missing");
    updateFlag(hub);
  }
  const { failure } = cacheFor(hub, on.connection.id);
  if (failure !== null && isCold(hub, on.connection.id, columns)) return failureBoard(api, failure);
  const board = boardOf(hub, api, on.connection, columns);
  const typed = query.trim();
  return typed === "" ? board : searched(hub, api, on, columns, typed, mine);
}
