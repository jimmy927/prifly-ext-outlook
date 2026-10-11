import { expect, test } from "bun:test";
import { categoryColour } from "../format";
import { boot, cleanUp, column } from "./support";

cleanUp();

test("the board has Unread, Inbox and Flagged, one card per conversation", async () => {
  const ext = await boot();
  const board = await ext.board();
  expect(board.columns.map((c) => c.name)).toEqual(["Unread", "Inbox", "Flagged"]);

  const unread = column(board, "Unread");
  expect(unread.map((card) => card.title)).toEqual(["Partner Center notice"]);
  expect(unread[0]?.detail).toBe("Dana Scully · 06:00");
  expect(unread[0]?.url).toBe("https://outlook.office.com/mail/id/m4");
  expect(unread[0]?.started).toBe(false);

  // Two conversations in the inbox, newest first, each once however many messages it has.
  const inbox = column(board, "Inbox");
  expect(inbox.map((card) => card.title)).toEqual(["Partner Center notice", "Contract for review"]);
  const contract = inbox[1];
  expect(contract?.detail).toBe("Bob Builder, Ada Lovelace · 9 Oct");
  expect(contract?.badges).toEqual({ comments: 2, attachments: 1 });
  // Everyone but the reader, the newest message's people first.
  expect(contract?.people).toEqual(["Bob Builder", "Dana Scully"]);
  expect(contract?.actions).toBeUndefined();
});

test("the flag's due day shows on the card, and is late once it has passed", async () => {
  const ext = await boot();
  const flagged = column(await ext.board(), "Flagged");
  expect(flagged).toHaveLength(1);
  expect(flagged[0]?.badges).toEqual({
    comments: 2,
    attachments: 2,
    due: "12 Oct",
    dueLate: false,
  });

  ext.clock.now = new Date("2026-10-13T08:00:00Z");
  const later = column(await ext.board(), "Flagged");
  expect(later[0]?.badges).toMatchObject({ due: "12 Oct", dueLate: true });
});

test("a category becomes a stripe colour by hashing its name", async () => {
  const ext = await boot({
    prepare: (graph) => {
      const second = graph.messages[1];
      if (second !== undefined) second["categories"] = ["Customers", "Urgent"];
    },
  });
  const inbox = column(await ext.board(), "Inbox");
  const card = inbox.find((each) => each.title === "Contract for review");
  expect(card?.stripes).toEqual([categoryColour("Customers"), categoryColour("Urgent")]);
  expect(categoryColour("customers")).toBe(categoryColour("Customers"));
  expect(categoryColour("Customers")).not.toBe("grey");
});

test("a column that finds nothing is still on the board", async () => {
  const ext = await boot({ settings: { columns: "Unread = unread | Nobody = from:nobody" } });
  const board = await ext.board();
  expect(board.columns.map((c) => c.name)).toEqual(["Unread", "Nobody"]);
  expect(column(board, "Nobody")).toEqual([]);
});

test("a column can be a KQL search, and a setting that does not read leaves the defaults", async () => {
  const ext = await boot({ settings: { columns: "From Dana = from:dana" } });
  const board = await ext.board();
  expect(board.columns.map((c) => c.name)).toEqual(["From Dana"]);
  expect(column(board, "From Dana").map((card) => card.title)).toEqual(["Partner Center notice"]);

  const broken = await boot({ settings: { columns: " | = | " } });
  expect((await broken.board()).columns.map((c) => c.name)).toEqual(["Unread", "Inbox", "Flagged"]);
});

test("a started conversation's card is tinted", async () => {
  const ext = await boot();
  await ext.extension.launch("mail", "convA", "", { again: false });
  await ext.extension.launched("mail", "convA", "session-1");
  const inbox = column(await ext.board(), "Inbox");
  expect(inbox.find((card) => card.key === "convA")?.started).toBe(true);
  expect(inbox.find((card) => card.key === "convB")?.started).toBe(false);
});

test("the filter searches with KQL once typing stops, and asks Graph once for a burst", async () => {
  const ext = await boot();
  await ext.board();
  const before = ext.graph.calls.length;
  const [first, second, last] = await Promise.all([
    ext.board("c"),
    ext.board("co"),
    ext.board("contract"),
  ]);
  const searches = ext.graph.calls.slice(before).filter((call) => call.includes("$search"));
  expect(searches).toHaveLength(1);
  expect(searches[0]).toContain("$search=%22contract%22");

  expect(last.columns).toEqual([{ id: "search", name: "Search: contract" }]);
  expect(last.items.map((card) => card.title)).toEqual(["Contract for review"]);
  // The earlier keystrokes were answered from what was in hand, with no round trip of their own.
  expect(first.columns.map((c) => c.name)).toEqual(["Unread", "Inbox", "Flagged"]);
  expect(second.items.every((card) => /co/i.test(`${card.title}${card.detail}`))).toBe(true);

  // The same search again, and the empty filter, come from memory.
  const settled = ext.graph.calls.length;
  await ext.board("contract");
  await ext.board("");
  expect(ext.graph.calls.length).toBe(settled);
});

test("a search that Graph refuses is a row with its text", async () => {
  const ext = await boot();
  await ext.board();
  ext.graph.fault = { kind: "throttle", status: 429, retryAfter: "1", times: 5 };
  const board = await ext.board("anything");
  expect(board.items).toHaveLength(1);
  expect(board.items[0]?.title).toContain("throttling requests (429)");
});

test("a mailbox that cannot be read is one row with the failure's text, never a throw", async () => {
  const ext = await boot({
    prepare: (graph) => {
      graph.fault = { kind: "throttle", status: 429, retryAfter: "1", times: 50 };
    },
  });
  const board = await ext.board();
  expect(board.items).toHaveLength(1);
  expect(board.items[0]?.key).toBe("row:error");
  expect(board.items[0]?.title).toBe(
    "Microsoft Graph is throttling requests (429); try again in 1 s.",
  );
  // Asked again a moment later it does not go back to Graph on every keystroke.
  const calls = ext.graph.calls.length;
  await ext.board("a");
  await ext.board("ab");
  expect(ext.graph.calls.length).toBe(calls);
});
