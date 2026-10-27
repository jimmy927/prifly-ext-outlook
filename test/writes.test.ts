import { afterAll, beforeAll, expect, test } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { NO_SEND } from "../api-tool";
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

test("every card has a one-sentence summary with no id or URL, a human content, and technical details", async () => {
  for (const [tool, args] of WRITES) {
    const { fake, graph } = setup();
    graph.messages.push({ ...DRAFT });
    await fake.callTool(tool, args);
    const card = fake.confirms[0];
    const label = `${tool} ${JSON.stringify(args)}`;
    expect(card?.summary, label).toMatch(/\.$/);
    expect(card?.summary, label).not.toMatch(/\b(m[1-4]|d1|draft1)\b|https?:|\/me\//);
    expect(card?.summary, label).not.toContain("\n");
    expect(card?.content?.trim() ?? "", label).not.toBe("");
    expect(card?.details.trim() ?? "", label).not.toBe("");
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
  expect(card?.summary).toBe("Send this mail to bob@example.com, dana@example.com.");
  expect(card?.content).toContain("To: bob@example.com");
  expect(card?.content).toContain("Cc: dana@example.com");
  expect(card?.content).toContain("Subject: Lunch?");
  expect(card?.content).toContain("Attachments: note.txt (1 KB)");
  expect(card?.content).toContain("Noon at the usual place?");
  expect(card?.details).toContain("POST /me/sendMail");
  expect(card?.details).not.toContain("Noon at the usual place?");
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
  expect(fake.confirms[0]?.summary).toBe("Send this reply to dana@example.com, bob@example.com.");
  expect(fake.confirms[0]?.content).toContain("Subject: RE: Partner Center notice");
  expect(fake.confirms[0]?.content).toContain("Thanks all.");
  expect(fake.confirms[0]?.content).not.toContain("To: Ada");
  expect(fake.confirms[0]?.details).toContain("/me/messages/m3/createReplyAll");
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
  expect(fake.confirms[0]?.summary).toBe("Forward this mail to dana@example.com.");
  expect(fake.confirms[0]?.content).toContain("Subject: FW: Contract for review");
  expect(fake.confirms[0]?.content).toContain("Attachments: note.txt");
  expect(fake.confirms[0]?.content).toContain("FYI");
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
  expect(fake.confirms[0]?.summary).toBe("Move 'Contract for review' to Inbox/Clients.");
  expect(fake.confirms[0]?.content).toContain("Contract for review");
  expect(fake.confirms[0]?.content).toContain("Bob");
  expect(fake.confirms[0]?.content).toContain("Moves to Inbox/Clients");
  expect(fake.confirms[0]?.content).not.toContain("m1");
  expect(fake.confirms[0]?.details).toContain("m1");
  expect(fake.confirms[0]?.details).toContain("f-clients");
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
  expect(fake.confirms[0]?.summary).toBe("Delete 'RE: Contract for review' to Deleted Items.");
  expect(fake.confirms[0]?.content).toContain("Moves to Deleted Items");
  expect(fake.confirms[0]?.details).toContain("m2");
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
  expect(fake.confirms[0]?.summary).toBe("Add categories to 'Partner Center notice': Red.");
  expect(fake.confirms[0]?.content).toContain("Categories: Partners → Partners, Red");
  expect(fake.confirms[0]?.details).toContain("m3");
  expect(graph.bodies.at(-1)?.body).toEqual({ categories: ["Partners", "Red"] });

  await fake.callTool("outlook_flag", { id: "m3", action: "complete" });
  expect(fake.confirms[1]?.summary).toBe("Mark 'Partner Center notice' as complete.");
  expect(fake.confirms[1]?.content).toContain("Flag: flagged → complete");
  expect(fake.confirms[1]?.details).toContain('"flagStatus":"complete"');
  expect(graph.bodies.at(-1)?.body).toEqual({ flag: { flagStatus: "complete" } });

  await fake.callTool("outlook_mark_read", { ids: ["m4"], read: true });
  expect(fake.confirms[2]?.summary).toBe("Mark 'RE: Partner Center notice' as read.");
  expect(fake.confirms[2]?.content).toContain("Unread → read");
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
  expect(card?.summary).toBe("Create an event.");
  expect(card?.content).toBe("Create an event in the mailbox of ada@example.com.");
  expect(card?.details).toContain("POST https://graph.microsoft.com/v1.0/me/events?%24select=id");
  expect(card?.details).toContain('"subject": "Lunch"');
});

test("outlook_api words a delete, and shows the body when it cannot say what the call does", async () => {
  const gone = setup();
  await gone.fake.callTool("outlook_api", { method: "DELETE", path: "/me/messages/m2" });
  expect(gone.fake.confirms[0]?.summary).toBe("Delete a message.");
  expect(gone.fake.confirms[0]?.details).toContain("DELETE https://graph.microsoft.com");

  const action = setup();
  await action.fake.callTool("outlook_api", {
    method: "POST",
    path: "/me/messages/m1/move",
    body: { destinationId: "archive" },
  });
  const card = action.fake.confirms[0];
  expect(card?.summary).toBe("Send a POST request to Microsoft Graph.");
  expect(card?.summary).not.toContain("http");
  expect(card?.content).toContain('"destinationId": "archive"');
  expect(card?.details).toContain("POST https://graph.microsoft.com/v1.0/me/messages/m1/move");
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

test("outlook_api refuses every call that sends mail, whatever the method, with no card and no request", async () => {
  const paths = [
    "/me/sendMail",
    "me/SENDMAIL",
    "/v1.0/me/sendMail",
    "https://graph.microsoft.com/v1.0/me/sendMail",
    "https://graph.microsoft.com/v1.0/users/bob@example.com/sendMail",
    "/users/bob@example.com/sendMail",
    "/v1.0/users/abc-123/messages/m1/send",
    "/me/messages/m1/send",
    "/me/messages/m1/Send/",
    "/me/messages/m1/send?$select=id",
    "/me/messages/m1/reply",
    "/me/messages/m1/replyAll",
    "/me/messages/m1/REPLYALL",
    "/me/messages/m1/forward",
    "/me/mailFolders/drafts/messages/d1/send",
    "https://graph.microsoft.com/beta/me/messages/m1/Forward?x=1",
    "/me/messages/m1/%73end",
    "/me/messages/m1/microsoft.graph.send",
    "/me/messages/m1/send/.",
    "/me/messages/m1/send/x/..",
  ];
  const { fake, connections } = setup({ choose: "send" });
  for (const path of paths) {
    for (const method of ["GET", "POST", "PATCH", "PUT", "DELETE"]) {
      const body = method === "GET" ? undefined : { comment: "hi" };
      const answer = await fake.callTool("outlook_api", { method, path, body });
      expect(answer, `${method} ${path}`).toBe(NO_SEND);
    }
  }
  expect(NO_SEND).toBe(
    "outlook_api does not send mail. Use outlook_send, outlook_reply, outlook_reply_all or outlook_forward: the reader chooses Send or Save as draft on their card.",
  );
  expect(connections.requests).toEqual([]);
  expect(fake.confirms).toEqual([]);
});

test("outlook_api refuses a $batch that carries a send", async () => {
  const { fake, connections } = setup();
  const answer = await fake.callTool("outlook_api", {
    method: "POST",
    path: "https://graph.microsoft.com/v1.0/$batch",
    body: { requests: [{ id: "1", method: "POST", url: "/me/messages/m1/send" }] },
  });
  expect(answer).toBe(NO_SEND);
  expect(connections.requests).toEqual([]);
  expect(fake.confirms).toEqual([]);
});

test("outlook_api still makes drafts and changes after the card, and says it never sends", async () => {
  for (const [method, path, body] of [
    ["POST", "/me/messages/m1/createReply", { comment: "Thanks" }],
    ["POST", "/me/messages/m3/createReplyAll", { comment: "Thanks" }],
    ["POST", "/me/messages/m1/createForward", { comment: "FYI" }],
    ["PATCH", "/me/messages/m1", { isRead: false }],
    ["POST", "/me/messages", { subject: "Draft" }],
  ] as const) {
    const { fake, graph } = setup({ choose: "send" });
    const answer = await fake.callTool("outlook_api", { method, path, body });
    expect(answer, path).toStartWith(
      `${method} https://graph.microsoft.com/v1.0${path} answered 20`,
    );
    expect(fake.confirms.length, path).toBe(1);
    expect(graph.sent, path).toEqual([]);
  }
  const { fake } = setup();
  const tool = fake.tools().find((each) => each.name === "outlook_api");
  expect(tool?.description).toContain("It never sends mail");
  expect(tool?.description).toContain("where the reader chooses Send or Save as draft");
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
