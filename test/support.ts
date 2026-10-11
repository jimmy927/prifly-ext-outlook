/**
 * What the board, poller and launch tests share: the extension on a fake api
 * and a fake Graph, with a clock and timers the test moves by hand.
 */

import { afterEach } from "bun:test";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { LaunchBoard, LaunchChoice, SettingValues } from "@prifly/extension-api";
import { makeExtension } from "../extension";
import { type FakeAccount, fakeApi, fakeConnections } from "./fake-api";
import { fakeGraph } from "./fake-graph";

// Times on the board are local; the tests read them in UTC.
process.env["TZ"] = "UTC";

export const ACCOUNT: FakeAccount = {
  id: "ms-1",
  provider: "microsoft",
  account: "ada@example.com",
  scopes: ["User.Read", "Mail.Read"],
};

export const SCOPES = { microsoft: ["Mail.Read", "Mail.ReadWrite"] };

/** Noon on the 11th of October 2026, when the fake mailbox's newest mail is a few hours old. */
export const NOON = new Date("2026-10-11T12:00:00Z");

const disposers: (() => void)[] = [];
const folders: string[] = [];

/** Each test file calls it once: what `boot` started is stopped and its folder removed after each test. */
export function cleanUp(): void {
  afterEach(async () => {
    for (const dispose of disposers.splice(0)) dispose();
    for (const dir of folders.splice(0)) await rm(dir, { recursive: true, force: true });
  });
}

/** Timers the test fires by hand. */
export function fakeTimers() {
  const running: { fn: () => void; ms: number }[] = [];
  return {
    running,
    every: (fn: () => void, ms: number) => {
      const entry = { fn, ms };
      running.push(entry);
      return () => {
        const at = running.indexOf(entry);
        if (at >= 0) running.splice(at, 1);
      };
    },
  };
}

export type BootOptions = {
  accounts?: FakeAccount[];
  settings?: Record<string, string | number>;
  /** Runs on the fake Graph before the extension is activated, to set the mailbox up. */
  prepare?: (graph: ReturnType<typeof fakeGraph>) => void;
  /** Replaces `api.connections.signIn`. */
  signIn?: (provider: string, scopes: string[]) => Promise<{ ok: false; message: string }>;
};

/** The extension, activated on a fake api; a poll round runs when the test asks (`poll()`). */
export async function boot(options: BootOptions = {}) {
  const folder = await mkdtemp(join(tmpdir(), "outlook-board-"));
  folders.push(folder);
  const graph = fakeGraph();
  options.prepare?.(graph);
  const accounts = options.accounts ?? [{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }];
  const connections = fakeConnections({ accounts, provider: graph.provider, scopes: SCOPES });
  const signIns: { scopes: string[]; connection: string | undefined }[] = [];
  const real = connections.connections.signIn;
  connections.connections.signIn = async (provider, scopes, signInOptions) => {
    signIns.push({ scopes, connection: signInOptions?.connection });
    return options.signIn === undefined
      ? real(provider, scopes, signInOptions)
      : options.signIn(provider, scopes);
  };
  const settings: SettingValues = {
    pollMinutes: 3,
    account: "",
    columns: "Unread = unread | Inbox = inbox | Flagged = flagged",
    ...options.settings,
  };
  const fake = fakeApi({ connections: connections.connections, settings, folder });
  const prompts: string[] = [];
  fake.api.prompt = async (_session, text) => {
    prompts.push(text);
    return { delivered: true };
  };
  let onSettings: ((values: SettingValues) => void) | null = null;
  fake.api.onSettings = (handler) => {
    onSettings = handler;
  };
  const clock = { now: NOON };
  const timers = fakeTimers();
  const extension = makeExtension({
    wait: async () => undefined,
    now: () => clock.now,
    debounceMs: 0,
    timers,
  });
  const dispose = extension.activate(fake.api);
  if (typeof dispose === "function") disposers.push(dispose);
  const poller = extension.poller();
  await poller?.tick();
  return {
    folder,
    graph,
    connections,
    fake,
    extension,
    clock,
    timers,
    prompts,
    signIns,
    settings,
    changeSettings: (next: Record<string, string | number>) => {
      Object.assign(settings, next);
      onSettings?.(settings);
    },
    dispose: () => (typeof dispose === "function" ? dispose() : undefined),
    poll: () => extension.poller()?.tick() ?? Promise.resolve(),
    board: async (query = ""): Promise<LaunchBoard> => {
      const answer = await extension.choices("mail", query);
      if (Array.isArray(answer)) throw new Error("expected a board, got a list");
      return answer;
    },
  };
}

export type Booted = Awaited<ReturnType<typeof boot>>;

/** The cards of one column of a board, by the column's name. */
export function column(board: LaunchBoard, name: string): LaunchChoice[] {
  const found = board.columns.find((each) => each.name === name);
  if (found === undefined)
    throw new Error(`No column ${name}: ${board.columns.map((c) => c.name)}`);
  return board.items.filter((item) => item.group === found.id);
}
