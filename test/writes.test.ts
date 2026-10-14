import { afterAll, beforeAll, expect, test } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { CANCELLED } from "../common";
import { mail, who } from "./fake-graph";
import { isWrite, setup } from "./harness";

let folder = "";
let note = "";

beforeAll(async () => {
  folder = await mkdtemp(join(tmpdir(), "outlook-writes-"));
  note = join(folder, "note.txt");
  await writeFile(note, "a small note");
});

afterAll(async () => {
  await rm(folder, { recursive: true, force: true });
});

/** A draft in the mailbox, for replace_draft_id. */
const DRAFT = mail({
  id: "d1",
  conversationId: "convA",
  subject: "RE: Contract for review",
  isDraft: true,
  from: who({ name: "Ada Lovelace", address: "ada@example.com" }),
  toRecipients: [who({ name: "Bob Builder", address: "bob@example.com" })],
  receivedDateTime: "2026-10-11T10:00:00Z",
});

/** Every write a tool can make, as a session would call it. */
const WRITES: [string, Record<string, unknown>][] = [
  ["outlook_send", { to: ["bob@example.com"], subject: "Lunch?", body: "Noon?" }],
  ["outlook_send", { to: "bob@example.com", subject: "File", body: "Here", attachments: [] }],
  ["outlook_reply", { id: "m1", body: "Thanks, signed." }],
  ["outlook_reply_all", { id: "m3", body: "Thanks all." }],
  ["outlook_forward", { id: "m1", to: ["dana@example.com"], body: "FYI" }],
  ["outlook_draft_reply", { id: "m1", body: "A draft" }],
  ["outlook_draft_reply", { id: "m1", body: "New text", replace_draft_id: "d1" }],
  ["outlook_move", { id: "m1", folder: "archive" }],
  ["outlook_delete", { ids: ["m1", "m2"] }],
  ["outlook_categories", { id: "m3", categories: ["Red"], mode: "add" }],
  ["outlook_flag", { id: "m1", action: "flag", due: "2026-10-20" }],
  ["outlook_mark_read", { id: "m4" }],
  ["outlook_api", { method: "POST", path: "/me/events", body: { subject: "Lunch" } }],
  ["outlook_api", { method: "PATCH", path: "/me/messages/m1", body: { isRead: false } }],
  ["outlook_api", { method: "DELETE", path: "https://graph.microsoft.com/v1.0/me/messages/m2" }],
];

test("every write shows one confirm card before Graph sees any write, then writes", async () => {
  for (const [tool, args] of WRITES) {
    const { fake, graph, writesBeforeConfirm } = setup();
    graph.messages.push({ ...DRAFT });
    const answer = await fake.callTool(tool, args);
    expect(fake.confirms.length, `${tool} ${answer}`).toBe(1);
    expect(writesBeforeConfirm, tool).toEqual([0]);
    expect(graph.calls.filter(isWrite).length, `${tool} ${answer}`).toBeGreaterThan(0);
  }
});

test("a cancelled card makes no write and answers Not done: you cancelled", async () => {
  for (const [tool, args] of WRITES) {
    const { fake, graph } = setup();
    graph.messages.push({ ...DRAFT });
    fake.answer.confirm = false;
    const answer = await fake.callTool(tool, args);
    expect(answer, tool).toBe(CANCELLED);
    expect(answer).toBe("Not done: you cancelled.");
    expect(fake.confirms.length, tool).toBe(1);
    expect(graph.calls.filter(isWrite), tool).toEqual([]);
    expect(graph.sent, tool).toEqual([]);
  }
});

test("a prifly without api.confirm writes nothing", async () => {
  const { fake, graph } = setup();
  delete fake.api.confirm;
  const answer = await fake.callTool("outlook_send", {
    to: ["bob@example.com"],
    subject: "x",
    body: "y",
  });
  expect(answer).toContain("Not done");
  expect(answer).toContain("Update prifly");
  expect(graph.calls.filter(isWrite)).toEqual([]);
});

test("outlook_send shows the recipients, subject, body and files, then sends exactly that", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_send", {
    to: ["bob@example.com"],
    cc: ["dana@example.com"],
    subject: "Lunch?",
    body: "Noon at the usual place?",
    attachments: [note],
  });
  const card = fake.confirms[0];
  expect(card?.title).toBe("Send an email");
  expect(card?.summary).toBe('Send "Lunch?" to bob@example.com, dana@example.com');
  expect(card?.details).toContain("To: bob@example.com");
  expect(card?.details).toContain("Cc: dana@example.com");
  expect(card?.details).toContain("Subject: Lunch?");
  expect(card?.details).toContain("Attachments: note.txt (1 KB)");
  expect(card?.details).toContain("Noon at the usual place?");
  expect(answer).toContain('Sent "Lunch?"');
  const sent = graph.sent[0] as Record<string, unknown>;
  expect(JSON.stringify(sent["toRecipients"])).toContain("bob@example.com");
  const files = sent["attachments"] as { name: string; contentBytes: string }[];
  expect(files[0]?.name).toBe("note.txt");
  expect(Buffer.from(files[0]?.contentBytes ?? "", "base64").toString()).toBe("a small note");
});

test("outlook_send refuses a missing file and a file over 3 MB before anything is asked or sent", async () => {
  const big = join(folder, "big.bin");
  await writeFile(big, new Uint8Array(3 * 1024 * 1024 + 1));
  for (const path of [big, join(folder, "missing.pdf")]) {
    const { fake, connections } = setup();
    const answer = await fake.callTool("outlook_send", {
      to: ["bob@example.com"],
      subject: "x",
      body: "y",
      attachments: [path],
    });
    expect(answer).toStartWith("Not done:");
    expect(fake.confirms).toEqual([]);
    expect(connections.requests).toEqual([]);
  }
});

test("outlook_reply_all shows everyone but the reader, and sends the draft Outlook made", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_reply_all", { id: "m3", body: "Thanks all." });
  expect(fake.confirms[0]?.summary).toBe(
    'Send "RE: Partner Center notice" to dana@example.com, bob@example.com',
  );
  expect(fake.confirms[0]?.details).toContain("Thanks all.");
  expect(fake.confirms[0]?.details).not.toContain("To: Ada");
  expect(graph.calls.filter(isWrite).map((call) => call.split("?")[0])).toEqual([
    "POST /me/messages/m3/createReplyAll",
    "POST /me/messages/draft1/send",
  ]);
  expect(answer).toContain("Sent");
});

test("a reply whose draft Outlook addresses differently is left in Drafts, not sent", async () => {
  const { fake, graph } = setup();
  graph.draftTo = [who({ name: "Eve", address: "eve@example.com" })];
  const answer = await fake.callTool("outlook_reply", { id: "m1", body: "Thanks" });
  expect(answer).toContain("Not sent: Outlook addressed it to eve@example.com");
  expect(answer).toContain("draft1");
  expect(graph.calls.some((call) => call.includes("/send"))).toBe(false);
  expect(graph.sent).toEqual([]);
});

test("outlook_forward adds the files to the draft, then sends it to the given address", async () => {
  const { fake, graph } = setup();
  await fake.callTool("outlook_forward", {
    id: "m1",
    to: ["dana@example.com"],
    body: "FYI",
    attachments: [note],
  });
  expect(fake.confirms[0]?.summary).toBe('Send "FW: Contract for review" to dana@example.com');
  expect(fake.confirms[0]?.details).toContain("Attachments: note.txt");
  expect(graph.calls.filter(isWrite).map((call) => call.split("?")[0])).toEqual([
    "POST /me/messages/m1/createForward",
    "POST /me/messages/draft1/attachments",
    "POST /me/messages/draft1/send",
  ]);
  expect(graph.attached.get("draft1")?.[0]?.["name"]).toBe("note.txt");
});

test("outlook_move finds a folder by name, and an unknown folder asks nothing", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_move", { id: "m1", folder: "clients" });
  expect(fake.confirms[0]?.summary).toBe("Move 1 message to Inbox/Clients");
  expect(answer).toContain("moved to Inbox/Clients");
  expect(graph.bodies.at(-1)?.body).toEqual({ destinationId: "f-clients" });

  const none = setup();
  expect(await none.fake.callTool("outlook_move", { id: "m1", folder: "Nowhere" })).toContain(
    "there is no folder Nowhere",
  );
  expect(none.fake.confirms).toEqual([]);
  expect(none.graph.calls.filter(isWrite)).toEqual([]);
});

test("outlook_delete moves to Deleted Items and says so", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_delete", { id: "m2" });
  expect(fake.confirms[0]?.summary).toBe("Move 1 message to Deleted Items");
  expect(graph.bodies.at(-1)).toEqual({
    call: "POST /me/messages/m2/move",
    body: { destinationId: "deleteditems" },
  });
  expect(answer).toContain("moved to Deleted Items");
  expect(answer).toContain("restored");
});

test("categories, flags and read state show the change on the card and PATCH exactly it", async () => {
  const { fake, graph } = setup();
  await fake.callTool("outlook_categories", { id: "m3", categories: ["Red"], mode: "add" });
  expect(fake.confirms[0]?.details).toContain("categories Partners → Partners, Red");
  expect(graph.bodies.at(-1)?.body).toEqual({ categories: ["Partners", "Red"] });

  await fake.callTool("outlook_flag", { id: "m3", action: "complete" });
  expect(fake.confirms[1]?.details).toContain("flag flagged → complete");
  expect(graph.bodies.at(-1)?.body).toEqual({ flag: { flagStatus: "complete" } });

  await fake.callTool("outlook_mark_read", { ids: ["m4"], read: true });
  expect(fake.confirms[2]?.summary).toBe("Mark 1 message as read");
  expect(graph.bodies.at(-1)?.body).toEqual({ isRead: true });
});

test("outlook_api shows the method, URL and body on the card", async () => {
  const { fake } = setup();
  await fake.callTool("outlook_api", {
    method: "post",
    path: "/me/events",
    query: { $select: "id" },
    body: { subject: "Lunch" },
  });
  const card = fake.confirms[0];
  expect(card?.summary).toBe(
    "POST https://graph.microsoft.com/v1.0/me/events?%24select=id as ada@example.com",
  );
  expect(card?.details).toContain('"subject": "Lunch"');
});

test("a GET through outlook_api runs at once, with no card", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_api", {
    path: "/me/messages/m1",
    query: { $select: "subject" },
  });
  expect(fake.confirms).toEqual([]);
  expect(answer).toContain(
    "GET https://graph.microsoft.com/v1.0/me/messages/m1?%24select=subject answered 200.",
  );
  expect(answer).toContain('"subject": "Contract for review"');
  expect(graph.calls.filter(isWrite)).toEqual([]);
});

test("outlook_api refuses any host but graph.microsoft.com, and paths outside /me, before any request", async () => {
  const refused = [
    "https://evil.example.com/v1.0/me/messages",
    "http://graph.microsoft.com/v1.0/me/messages",
    "https://graph.microsoft.com.evil.example.com/v1.0/me",
    "https://user:pw@graph.microsoft.com/v1.0/me",
    "https://graph.microsoft.com:8443/v1.0/me",
    "/users/bob@example.com/messages",
    "/me/../users/bob/messages",
    "/v1.0/meeting",
  ];
  for (const path of refused) {
    for (const method of ["GET", "POST"]) {
      const { fake, connections } = setup();
      const answer = await fake.callTool("outlook_api", {
        method,
        path,
        body: method === "GET" ? undefined : {},
      });
      expect(answer, path).toStartWith("Not done:");
      expect(connections.requests, path).toEqual([]);
      expect(fake.confirms, path).toEqual([]);
    }
  }
});
