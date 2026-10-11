import { expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { ADA, body, DANA, mail, who } from "./fake-graph";
import { ACCOUNT, type Booted, boot, cleanUp } from "./support";

cleanUp();

/** A new message in the fake mailbox. */
function arrive(ext: Booted, fields: Record<string, unknown>) {
  ext.graph.messages.push(
    mail({
      id: `n${ext.graph.messages.length}`,
      conversationId: "convNew",
      subject: "Something new",
      from: who(DANA),
      toRecipients: [who(ADA)],
      receivedDateTime: "2026-10-11T11:00:00Z",
      isRead: false,
      bodyPreview: "hello",
      body: body("hello"),
      ...fields,
    }),
  );
}

const deltaCalls = (ext: Booted) => ext.graph.calls.filter((call) => call.includes("/delta"));

test("the first round reads the recent inbox, lights the launcher and keeps the delta link", async () => {
  const ext = await boot();
  expect(ext.fake.flags.get("mail")).toEqual({ tone: "info", hint: "1 unread" });
  const [seed] = deltaCalls(ext);
  expect(decodeURIComponent(seed ?? "")).toContain(
    "$filter=receivedDateTime ge 2026-09-11T12:00:00.000Z",
  );
  // The delta pages through to its end (two pages in the fake).
  expect(deltaCalls(ext)).toHaveLength(2);

  const state = JSON.parse(await readFile(join(ext.folder, "state.json"), "utf8"));
  expect(state.deltas["ms-1"]).toContain("$deltatoken=T4");
});

test("a delta that adds an unread mail raises the unread count on the launcher", async () => {
  const ext = await boot();
  arrive(ext, {});
  await ext.poll();
  expect(ext.fake.flags.get("mail")).toEqual({ tone: "info", hint: "2 unread" });
  // The second round followed the saved link rather than reading the inbox again.
  expect(deltaCalls(ext).at(-1)).toContain("$deltatoken=T4");

  const board = await ext.board();
  expect(board.items.some((card) => card.title === "Something new")).toBe(true);
});

test("a quiet round only looks at flagged mail, and puts the launcher out when nothing is unread", async () => {
  const ext = await boot();
  const before = ext.graph.calls.length;
  await ext.poll();
  const round = ext.graph.calls.slice(before);
  expect(round.map((call) => call.split("?")[0])).toEqual([
    "GET /me/mailFolders/inbox/messages/delta",
    "GET /me/messages",
  ]);
  expect(round[1]).toContain("flagStatus");

  const read = ext.graph.messages.find((m) => m["id"] === "m4");
  if (read !== undefined) read["isRead"] = true;
  arrive(ext, { isRead: true });
  await ext.poll();
  expect(ext.fake.flags.get("mail")).toBeNull();
});

test("a reply in a conversation a session came from is a chip on that session, with no prompt", async () => {
  const ext = await boot();
  await ext.extension.launch("mail", "convB", "", { again: false });
  await ext.extension.launched("mail", "convB", "session-1");
  arrive(ext, { conversationId: "convB", subject: "RE: Partner Center notice" });
  await ext.poll();

  const shown = ext.fake.shown.at(-1);
  expect(Object.keys(shown?.bySession ?? {})).toEqual(["session-1"]);
  const chip = shown?.bySession["session-1"]?.[0];
  expect(chip?.label).toBe("New reply from Dana Scully");
  expect(chip?.key).toBe("convB");
  expect(chip?.launcher).toBe("mail");
  expect(chip?.url).toBe("https://outlook.office.com/mail/id/n4");
  expect(shown?.statusBar).toEqual([]);
  expect(ext.prompts).toEqual([]);

  // Reading the conversation inside prifly clears it.
  await ext.extension.open("mail", "convB");
  expect(ext.fake.shown.at(-1)?.bySession).toEqual({});
});

test("a mail in a conversation nobody started from, or sent by the reader, is no chip", async () => {
  const ext = await boot();
  await ext.extension.launch("mail", "convB", "", { again: false });
  await ext.extension.launched("mail", "convB", "session-1");
  arrive(ext, { conversationId: "convA" });
  arrive(ext, { conversationId: "convB", from: who(ADA), isRead: false });
  await ext.poll();
  expect(ext.fake.shown.at(-1)?.bySession).toEqual({});
});

test("an expired delta link (410) starts over from the recent inbox", async () => {
  const ext = await boot();
  ext.graph.expiredDelta = true;
  arrive(ext, {});
  await ext.poll();
  const calls = deltaCalls(ext);
  const expired = calls.findIndex((call) => call.includes("$deltatoken=T4"));
  expect(expired).toBeGreaterThan(1);
  // Right after the 410 it read the inbox again with a filter and no token.
  const retry = calls[expired + 1] ?? "";
  expect(retry).toContain("receivedDateTime");
  expect(retry).not.toContain("deltatoken");
  expect(ext.fake.flags.get("mail")).toEqual({ tone: "info", hint: "2 unread" });
  expect(ext.extension.hub()?.state.deltas["ms-1"]).toContain("$deltatoken=T5");
});

test("the timer fires a round every pollMinutes, a new setting restarts it, the disposer stops it", async () => {
  const ext = await boot();
  expect(ext.timers.running.map((timer) => timer.ms)).toEqual([180_000]);

  ext.changeSettings({ pollMinutes: 10 });
  expect(ext.timers.running.map((timer) => timer.ms)).toEqual([600_000]);

  arrive(ext, {});
  ext.timers.running[0]?.fn();
  await ext.poll();
  expect(ext.fake.flags.get("mail")?.hint).toBe("2 unread");

  ext.dispose();
  expect(ext.timers.running).toEqual([]);
  ext.changeSettings({ pollMinutes: 5 });
  expect(ext.timers.running).toEqual([]);
});

test("the poller never opens the browser and skips what it cannot read", async () => {
  const ext = await boot({
    accounts: [
      { ...ACCOUNT, scopes: ["User.Read"] },
      { id: "ms-2", provider: "microsoft", account: "bob@example.com", scopes: ["Mail.Read"] },
    ],
  });
  await ext.poll();
  expect(ext.signIns).toEqual([]);
  // Only the connection holding Mail.Read was read.
  expect(ext.connections.requests.length).toBeGreaterThan(0);
  expect(ext.connections.endpoint.issued("ms-1")).toBe(0);
  expect(ext.connections.endpoint.issued("ms-2")).toBeGreaterThan(0);
});

test("a failing round is logged and the next one tries again", async () => {
  const ext = await boot();
  ext.graph.fault = { kind: "throttle", status: 429, retryAfter: "1", times: 50 };
  await ext.poll();
  expect(ext.extension.hub()?.caches.get("ms-1")?.failure).toContain("throttling");
  ext.graph.fault = null;
  arrive(ext, {});
  await ext.poll();
  expect(ext.fake.flags.get("mail")).toEqual({ tone: "info", hint: "2 unread" });
  expect(ext.extension.hub()?.caches.get("ms-1")?.failure).toBeNull();
});
