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
  folder = await mkdtemp(join(tmpdir(), "outlook-choose-"));
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

/** A Graph call that sends mail at once. */
const SENDS = /^POST \/me\/(sendMail|messages\/[^/]+\/(send|reply|replyAll|forward))(\?|$)/;

/** The five mail tools (six ways in): arguments, the card's title, and what the name does without the two-way card. */
const COMPOSE: [string, Record<string, unknown>, string, "draft" | "send"][] = [
  [
    "outlook_send",
    { to: ["bob@example.com"], subject: "Lunch?", body: "Noon?" },
    "Write an email",
    "send",
  ],
  ["outlook_reply", { id: "m1", body: "Thanks, signed." }, "Reply", "send"],
  ["outlook_reply_all", { id: "m3", body: "Thanks all." }, "Reply", "send"],
  [
    "outlook_forward",
    { id: "m1", to: ["dana@example.com"], body: "FYI" },
    "Forward an email",
    "send",
  ],
  ["outlook_draft_reply", { id: "m1", body: "A draft" }, "Reply", "draft"],
  ["outlook_draft_reply", { body: "New text", replace_draft_id: "d1" }, "Reply", "draft"],
];

function world(choose?: string | null) {
  const w = setup(choose === undefined ? {} : { choose });
  w.graph.messages.push({ ...DRAFT });
  const writes = () => w.graph.calls.filter(isWrite).map((call) => call.split("?")[0] ?? "");
  return { ...w, writes };
}

for (const [tool, args, title, own] of COMPOSE) {
  const label = `${tool} ${JSON.stringify(args)}`;

  test(`${label}: "draft" on the two-way card saves it in Drafts and sends nothing`, async () => {
    const w = world("draft");
    const answer = await w.fake.callTool(tool, args);
    expect(answer).toStartWith("The reader saved it in Drafts: ");
    expect(answer).toContain("is in Drafts, not sent.");
    expect(w.chooses.length).toBe(1);
    expect(w.chooses[0]?.writesBefore).toBe(0);
    expect(w.fake.confirms).toEqual([]);
    expect(w.chooses[0]?.spec.title).toBe(title);
    expect(w.chooses[0]?.spec.summary).toEndWith(". Save it in Drafts, or send it now.");
    expect(w.chooses[0]?.spec.mail.draft).toBe(false);
    expect(w.writes().length).toBeGreaterThan(0);
    expect(w.writes().filter((call) => SENDS.test(call))).toEqual([]);
    expect(w.graph.sent).toEqual([]);
  });

  test(`${label}: "send" on the two-way card sends it`, async () => {
    const w = world("send");
    const answer = await w.fake.callTool(tool, args);
    expect(answer).toStartWith("The reader sent it: Sent ");
    expect(w.chooses.length).toBe(1);
    expect(w.chooses[0]?.writesBefore).toBe(0);
    expect(w.fake.confirms).toEqual([]);
    expect(w.graph.sent.length).toBe(1);
  });

  test(`${label}: Cancel on the two-way card writes nothing`, async () => {
    const w = world(null);
    expect(await w.fake.callTool(tool, args)).toBe(CANCELLED);
    expect(w.chooses.length).toBe(1);
    expect(w.fake.confirms).toEqual([]);
    expect(w.writes()).toEqual([]);
    expect(w.graph.sent).toEqual([]);
  });

  test(`${label}: without "choose-mail" the confirm card does what the name says (${own})`, async () => {
    const w = world();
    const answer = await w.fake.callTool(tool, args);
    expect(w.chooses).toEqual([]);
    expect(w.fake.confirms.length).toBe(1);
    expect(answer).not.toContain("The reader");
    expect(w.graph.sent.length).toBe(own === "send" ? 1 : 0);
  });
}

test("an answer the reader did not give on the card is a Cancel", async () => {
  for (const odd of ["Send", "yes", "", "sent"]) {
    const w = world(odd);
    expect(await w.fake.callTool("outlook_send", COMPOSE[0]?.[1] ?? {})).toBe(CANCELLED);
    expect(w.writes()).toEqual([]);
  }
});

test("outlook_send's card shows the mail and its file; saved, it is a new draft with the file", async () => {
  const w = world("draft");
  const answer = await w.fake.callTool("outlook_send", {
    to: ["bob@example.com"],
    cc: ["dana@example.com"],
    subject: "Notes",
    body: "Here they are.",
    attachments: [note],
  });
  const spec = w.chooses[0]?.spec;
  expect(spec?.summary).toBe(
    "“Notes” to bob@example.com, dana@example.com, with 1 attachment. Save it in Drafts, or send it now.",
  );
  expect(spec?.mail).toMatchObject({
    to: "bob@example.com",
    cc: "dana@example.com",
    subject: "Notes",
    text: "Here they are.",
  });
  expect(spec?.mail.bcc).toBeUndefined();
  const file = spec?.mail.attachments?.[0];
  expect(file?.filename).toBe("note.txt");
  expect(file?.mimeType).toBe("text/plain");
  expect(new TextDecoder().decode(file?.bytes)).toBe("a small note");
  expect(w.writes()).toEqual(["POST /me/messages"]);
  const saved = w.graph.bodies.at(-1)?.body ?? {};
  expect(JSON.stringify(saved["toRecipients"])).toContain("bob@example.com");
  const files = saved["attachments"] as { name: string; contentBytes: string }[];
  expect(Buffer.from(files[0]?.contentBytes ?? "", "base64").toString()).toBe("a small note");
  expect(answer).toBe(
    'The reader saved it in Drafts: "Notes" to bob@example.com, dana@example.com is in Drafts, not sent. Draft id: draft1\nOpen it in Outlook: https://outlook.office.com/mail/id/draft1',
  );
});

test("outlook_send with Send now goes out through sendMail", async () => {
  const w = world("send");
  const answer = await w.fake.callTool("outlook_send", COMPOSE[0]?.[1] ?? {});
  expect(w.writes()).toEqual(["POST /me/sendMail"]);
  expect(answer).toBe(
    'The reader sent it: Sent "Lunch?" to bob@example.com from ada@example.com. A copy is in Sent Items.',
  );
});

test("a reply's card names the people in plain Unicode; saved, it stays a threaded draft", async () => {
  const w = world("draft");
  w.graph.messages.push(
    mail({
      id: "m5",
      conversationId: "convC",
      subject: "Fika på fredag?",
      from: who({ name: "Åsa Öberg", address: "asa@example.com" }),
      toRecipients: [who({ name: "Ada Lovelace", address: "ada@example.com" })],
      receivedDateTime: "2026-10-12T08:00:00Z",
    }),
  );
  await w.fake.callTool("outlook_reply", { id: "m5", body: "Gärna!" });
  const spec = w.chooses[0]?.spec;
  expect(spec?.summary).toBe(
    "“RE: Fika på fredag?” to Åsa Öberg. Save it in Drafts, or send it now.",
  );
  expect(spec?.mail.to).toBe("Åsa Öberg <asa@example.com>");
  expect(spec?.mail.text).toStartWith("Gärna!");
  expect(w.writes()).toEqual(["POST /me/messages/m5/createReply"]);
});

test("a forward saved as a draft gets its files, and is not sent", async () => {
  const w = world("draft");
  const answer = await w.fake.callTool("outlook_forward", {
    id: "m1",
    to: ["dana@example.com"],
    body: "FYI",
    attachments: [note],
  });
  expect(w.chooses[0]?.spec.summary).toBe(
    "“FW: Contract for review” to dana@example.com, with 1 attachment. Save it in Drafts, or send it now.",
  );
  expect(w.writes()).toEqual([
    "POST /me/messages/m1/createForward",
    "POST /me/messages/draft1/attachments",
  ]);
  expect(answer).toStartWith('The reader saved it in Drafts: "FW: Contract for review"');
});

test("outlook_draft_reply with Send now sends the reply Outlook drafted, once", async () => {
  const w = world("send");
  await w.fake.callTool("outlook_draft_reply", { id: "m3", body: "Thanks all.", all: true });
  expect(w.writes()).toEqual([
    "POST /me/messages/m3/createReplyAll",
    "POST /me/messages/draft1/send",
  ]);
  expect(w.graph.sent.length).toBe(1);
});

test("a reply Outlook addresses differently is not sent on Send now, and flagged when saved", async () => {
  const sent = world("send");
  sent.graph.draftTo = [who({ name: "Eve", address: "eve@example.com" })];
  const answer = await sent.fake.callTool("outlook_reply", { id: "m1", body: "Thanks" });
  expect(answer).toStartWith("Not sent: Outlook addressed it to eve@example.com");
  expect(sent.graph.sent).toEqual([]);

  const saved = world("draft");
  saved.graph.draftTo = [who({ name: "Eve", address: "eve@example.com" })];
  const kept = await saved.fake.callTool("outlook_reply", { id: "m1", body: "Thanks" });
  expect(kept).toContain("not to the bob@example.com on the card: check it before it is sent.");
  expect(saved.graph.sent).toEqual([]);
});

test("replace_draft_id: the card shows the draft as it will be; Send now updates it, then sends it", async () => {
  const w = world("send");
  const answer = await w.fake.callTool("outlook_draft_reply", {
    body: "New text",
    replace_draft_id: "d1",
  });
  expect(w.chooses[0]?.spec.summary).toBe(
    "“RE: Contract for review” to Bob Builder. Save it in Drafts, or send it now.",
  );
  expect(w.chooses[0]?.spec.mail).toMatchObject({
    to: "Bob Builder <bob@example.com>",
    text: "New text",
  });
  expect(w.writes()).toEqual(["PATCH /me/messages/d1", "POST /me/messages/d1/send"]);
  expect(answer).toBe(
    'The reader sent it: Sent "RE: Contract for review" to bob@example.com from ada@example.com. A copy is in Sent Items.',
  );
});

test("the two-way card is asked bound to prifly's api", async () => {
  const w = world();
  let bound = false;
  const api = w.fake.api;
  Object.assign(api, {
    features: [...(api.features ?? []), "choose-mail"],
    chooseMail(this: unknown) {
      bound = this === api;
      return Promise.resolve(null);
    },
  });
  expect(await w.fake.callTool("outlook_send", COMPOSE[0]?.[1] ?? {})).toBe(CANCELLED);
  expect(bound).toBe(true);
});

test("every mail tool's description says the reader decides, and not to send another way", () => {
  const w = world();
  for (const name of new Set(COMPOSE.map(([tool]) => tool))) {
    const description = w.fake.tools().find((each) => each.name === name)?.description ?? "";
    expect(description, name).toContain("chooses there to save it in Drafts or send it now");
    expect(description, name).toContain("Never try another way to send mail.");
  }
});
