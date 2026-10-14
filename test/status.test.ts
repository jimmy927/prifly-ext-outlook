import { expect, test } from "bun:test";
import { makeExtension } from "../extension";
import { STATUS_EVERY_MS } from "../status";
import { type FakeAccount, fakeApi, fakeConnections } from "./fake-api";
import { fakeGraph } from "./fake-graph";
import { ACCOUNT, SCOPES } from "./harness";

/** The extension activated on a fake api, with the status timer in the test's hands. */
function boot(accounts: FakeAccount[] = [{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }]) {
  const graph = fakeGraph();
  const connections = fakeConnections({ accounts, provider: graph.provider, scopes: SCOPES });
  const signIns: string[][] = [];
  connections.connections.signIn = async (_provider, scopes) => {
    signIns.push(scopes);
    return { ok: false, message: "not in a test" };
  };
  const fake = fakeApi({ connections: connections.connections });
  const timers: { fn: () => void; ms: number; stopped: boolean }[] = [];
  const extension = makeExtension({
    wait: async () => undefined,
    every: (fn, ms) => {
      const timer = { fn, ms, stopped: false };
      timers.push(timer);
      return () => {
        timer.stopped = true;
      };
    },
  });
  const dispose = extension.activate(fake.api);
  return { graph, connections, fake, extension, timers, signIns, dispose };
}

test("the status line says who is connected and how much is unread, once at start", async () => {
  const { fake, extension, graph } = boot();
  await extension.refreshStatus();
  expect(fake.statuses).toEqual(["Connected as ada@example.com · 1 unread"]);
  expect(graph.calls).toEqual(["GET /me/mailFolders/inbox?$select=unreadItemCount"]);
});

test("it is refreshed every 15 minutes, and no more after the extension stops", async () => {
  const { fake, extension, graph, timers, dispose } = boot();
  await extension.refreshStatus();
  expect(timers.map((t) => t.ms)).toEqual([STATUS_EVERY_MS]);
  expect(STATUS_EVERY_MS).toBe(15 * 60 * 1000);
  for (const each of graph.messages) each["isRead"] = false;
  timers[0]?.fn();
  await extension.refreshStatus();
  expect(fake.statuses.at(-1)).toBe("Connected as ada@example.com · 4 unread");
  if (typeof dispose === "function") dispose();
  expect(timers[0]?.stopped).toBe(true);
});

test("without a connection, or with an ended grant, it says it is not connected", async () => {
  const none = boot([]);
  await none.extension.refreshStatus();
  expect(none.fake.statuses).toEqual([
    "Not connected: No Microsoft account is connected: connect one in prifly's Vault → Connections.",
  ]);

  const revoked = boot();
  await revoked.extension.refreshStatus();
  revoked.connections.endpoint.revoke("ms-1");
  await revoked.extension.refreshStatus();
  expect(revoked.fake.statuses.at(-1)).toBe(
    "Not connected: The Outlook grant for ada@example.com ended: sign in again in Vault → Connections.",
  );
});

test("a connection without mail access says so and never opens the browser", async () => {
  const { fake, extension, signIns, graph } = boot([{ ...ACCOUNT, scopes: ["User.Read"] }]);
  await extension.refreshStatus();
  expect(fake.statuses.at(-1)).toContain(
    "Connected as ada@example.com · mail access not approved yet",
  );
  expect(signIns).toEqual([]);
  expect(graph.calls).toEqual([]);
});
