/**
 * What the extension keeps between runs, in `state.json` in its folder: the
 * delta link of each connection's inbox (not a secret), which account the
 * board shows, and the conversations sessions were started from. Read
 * tolerantly: a file that does not read is an empty state, never an error.
 */

import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { record, text } from "./model";

/** A reply that came to a started conversation and has not been looked at yet. */
export type Reply = { from: string; at: string; subject: string; url: string };

/** A conversation a session was started from. */
export type Started = {
  session: string;
  connection: string;
  /** ISO time of the newest message the session knows of. */
  seen: string;
  reply: Reply | null;
};

export type State = {
  /** The connection the reader switched the board to; empty: the `account` setting decides. */
  account: string;
  /** Inbox delta links, by connection id. */
  deltas: Record<string, string>;
  /** By conversation id. */
  started: Record<string, Started>;
};

export const STATE_FILE = "state.json";

export function emptyState(): State {
  return { account: "", deltas: {}, started: {} };
}

function toReply(value: unknown): Reply | null {
  const r = record(value);
  if (text(r["at"]) === "") return null;
  return {
    from: text(r["from"]),
    at: text(r["at"]),
    subject: text(r["subject"]),
    url: text(r["url"]),
  };
}

function toStarted(value: unknown): Started | null {
  const r = record(value);
  if (text(r["session"]) === "") return null;
  return {
    session: text(r["session"]),
    connection: text(r["connection"]),
    seen: text(r["seen"]),
    reply: toReply(r["reply"]),
  };
}

export function parseState(raw: unknown): State {
  const r = record(raw);
  const state = emptyState();
  state.account = text(r["account"]);
  for (const [id, link] of Object.entries(record(r["deltas"]))) {
    if (typeof link === "string" && link !== "") state.deltas[id] = link;
  }
  for (const [key, value] of Object.entries(record(r["started"]))) {
    const started = toStarted(value);
    if (started !== null) state.started[key] = started;
  }
  return state;
}

export async function loadState(folder: string): Promise<State> {
  try {
    return parseState(JSON.parse(await readFile(join(folder, STATE_FILE), "utf8")));
  } catch {
    return emptyState();
  }
}

/** Writes the state; null when it landed, else why not. */
export async function saveState(folder: string, state: State): Promise<string | null> {
  try {
    await mkdir(folder, { recursive: true });
    const path = join(folder, STATE_FILE);
    // Written beside and renamed over, so a crash never leaves half a file.
    await writeFile(`${path}.tmp`, JSON.stringify(state, null, 2));
    await rename(`${path}.tmp`, path);
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}
