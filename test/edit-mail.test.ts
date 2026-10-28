import { expect, test } from "bun:test";
import { CANCELLED } from "../common";
import { mail, who } from "./fake-graph";
import { isWrite, setup } from "./harness";

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

const SEND = { to: ["bob@example.com"], subject: "Lunch?", body: "Noon?" };

function world(options: Parameters<typeof setup>[0]) {
  const w = setup(options);
  w.graph.messages.push({ ...DRAFT });
  const writes = () => w.graph.calls.filter(isWrite).map((call) => call.split("?")[0] ?? "");
  /** The JSON body of the one request to `call`. */
  const bodyOf = (call: string) => w.graph.bodies.find((each) => each.call === call)?.body;
  return { ...w, writes, bodyOf };
}

/** The editable card's answer: the mail as shown, with these changes. */
const edits =
  (choice: string, change: { subject?: string; text?: (shown: string) => string }) =>
  (shown: { subject: string; text: string }) => ({
    choice,
    subject: change.subject ?? shown.subject,
    text: change.text === undefined ? shown.text : change.text(shown.text),
  });

test("outlook_send: Send now goes out with the edited subject and text, and the answer says so", async () => {
  const w = world({
    edit: edits("send", { subject: "Lunch tomorrow?", text: () => "Noon works?" }),
  });
  const answer = await w.fake.callTool("outlook_send", SEND);
  expect(w.chooses[0]?.spec.mail).toMatchObject({ subject: "Lunch?", text: "Noon?" });
  expect(w.writes()).toEqual(["POST /me/sendMail"]);
  const sent = w.graph.sent[0] ?? {};
  expect(sent["subject"]).toBe("Lunch tomorrow?");
  expect(sent["body"]).toEqual({ contentType: "Text", content: "Noon works?" });
  expect(answer).toBe(
    'The reader sent it: Sent "Lunch tomorrow?" to bob@example.com from ada@example.com. A copy is in Sent Items.\n' +
      'The reader edited the mail first. Its subject now: "Lunch tomorrow?". Its text now:\nNoon works?',
  );
});

test("outlook_send: Save as draft keeps the edited text, and a text-only edit leaves the subject", async () => {
  const w = world({ edit: edits("draft", { text: (text) => `${text} Or 1pm.` }) });
  const answer = await w.fake.callTool("outlook_send", SEND);
  expect(w.writes()).toEqual(["POST /me/messages"]);
  expect(w.bodyOf("POST /me/messages")).toMatchObject({
    subject: "Lunch?",
    body: { contentType: "Text", content: "Noon? Or 1pm." },
  });
  expect(w.graph.sent).toEqual([]);
  expect(answer).toStartWith('The reader saved it in Drafts: "Lunch?" to bob@example.com');
  expect(answer).toEndWith('Its subject now: "Lunch?". Its text now:\nNoon? Or 1pm.');
});

test("outlook_send: an unedited mail is written and answered exactly as on the plain two-way card", async () => {
  const edited = world({ edit: edits("send", {}) });
  const plain = world({ choose: "send" });
  const answer = await edited.fake.callTool("outlook_send", SEND);
  expect(answer).toBe(await plain.fake.callTool("outlook_send", SEND));
  expect(answer).not.toContain("edited");
  expect(edited.graph.calls).toEqual(plain.graph.calls);
  expect(edited.graph.sent).toEqual(plain.graph.sent);
});

test("outlook_reply: the edited text is the comment, the edited subject is patched in, then it is sent", async () => {
  const w = world({
    edit: edits("send", {
      subject: "RE: Contract, signed",
      text: (text) => text.replace("Thanks, signed.", "Signed, thanks!"),
    }),
  });
  const answer = await w.fake.callTool("outlook_reply", { id: "m1", body: "Thanks, signed." });
  expect(w.chooses[0]?.spec.mail.text).toStartWith(
    "Thanks, signed.\n\n(Below it: the quoted mail from",
  );
  expect(w.writes()).toEqual([
    "POST /me/messages/m1/createReply",
    "PATCH /me/messages/draft1",
    "POST /me/messages/draft1/send",
  ]);
  // The card's own line about the quoted mail is not part of the comment.
  expect(w.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "Signed, thanks!" });
  expect(w.bodyOf("PATCH /me/messages/draft1")).toEqual({ subject: "RE: Contract, signed" });
  expect(w.graph.sent.length).toBe(1);
  expect(answer).toStartWith('The reader sent it: Sent "RE: Contract, signed" to bob@example.com');
  expect(answer).toEndWith(
    'Its subject now: "RE: Contract, signed". Its text now:\nSigned, thanks!',
  );
});

test("outlook_reply: a text-only edit patches nothing, an unedited reply is as today", async () => {
  const text = world({
    edit: edits("draft", { text: (shown) => shown.replace("Thanks", "Many thanks") }),
  });
  await text.fake.callTool("outlook_reply", { id: "m1", body: "Thanks, signed." });
  expect(text.writes()).toEqual(["POST /me/messages/m1/createReply"]);
  expect(text.bodyOf("POST /me/messages/m1/createReply")).toEqual({
    comment: "Many thanks, signed.",
  });

  const plain = world({ edit: edits("draft", {}) });
  const answer = await plain.fake.callTool("outlook_reply", { id: "m1", body: "Thanks, signed." });
  expect(plain.writes()).toEqual(["POST /me/messages/m1/createReply"]);
  expect(plain.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "Thanks, signed." });
  expect(answer).not.toContain("edited");
});

test("outlook_reply: a reader who removed the quoted-mail line gets their text as it stands", async () => {
  const w = world({
    edit: () => ({ choice: "draft", subject: "RE: Contract for review", text: "Just this" }),
  });
  const answer = await w.fake.callTool("outlook_reply", { id: "m1", body: "Thanks, signed." });
  expect(w.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "Just this" });
  expect(answer).toEndWith("Its text now:\nJust this");
});

test("outlook_forward: the edited subject is patched into the forward draft, files still added", async () => {
  const w = world({ edit: edits("draft", { subject: "FW: For Dana" }) });
  const answer = await w.fake.callTool("outlook_forward", {
    id: "m1",
    to: ["dana@example.com"],
    body: "FYI",
  });
  expect(w.writes()).toEqual(["POST /me/messages/m1/createForward", "PATCH /me/messages/draft1"]);
  expect(answer).toContain('"FW: For Dana" to dana@example.com is in Drafts, not sent.');
});

test("outlook_reply: an edited reply Outlook addresses differently is still not sent", async () => {
  const w = world({ edit: edits("send", { subject: "RE: Other" }) });
  w.graph.draftTo = [who({ name: "Eve", address: "eve@example.com" })];
  const answer = await w.fake.callTool("outlook_reply", { id: "m1", body: "Thanks" });
  expect(answer).toStartWith("Not sent: Outlook addressed it to eve@example.com");
  expect(w.graph.sent).toEqual([]);
  expect(w.writes()).toEqual(["POST /me/messages/m1/createReply"]);
});

test("outlook_draft_reply: the edited text and subject are written to the reply draft", async () => {
  const w = world({ edit: edits("draft", { subject: "RE: Short", text: () => "Edited" }) });
  const answer = await w.fake.callTool("outlook_draft_reply", { id: "m1", body: "A draft" });
  expect(w.writes()).toEqual(["POST /me/messages/m1/createReply", "PATCH /me/messages/draft1"]);
  expect(w.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "Edited" });
  expect(answer).toStartWith('The reader saved it in Drafts: "RE: Short"');
  expect(answer).toContain("Its text now:\nEdited");
});

test("outlook_draft_reply replace_draft_id: one PATCH carries the edited subject and text, then it is sent", async () => {
  const w = world({
    edit: edits("send", { subject: "RE: Contract, final", text: () => "Final text" }),
  });
  const answer = await w.fake.callTool("outlook_draft_reply", {
    body: "New text",
    replace_draft_id: "d1",
  });
  expect(w.writes()).toEqual(["PATCH /me/messages/d1", "POST /me/messages/d1/send"]);
  expect(w.bodyOf("PATCH /me/messages/d1")).toEqual({
    subject: "RE: Contract, final",
    body: { contentType: "text", content: "Final text" },
  });
  expect(answer).toStartWith('The reader sent it: Sent "RE: Contract, final" to bob@example.com');
  expect(answer).toEndWith("Its text now:\nFinal text");
});

test("outlook_draft_reply replace_draft_id: unedited, the PATCH is the body alone", async () => {
  const w = world({ edit: edits("draft", {}) });
  const answer = await w.fake.callTool("outlook_draft_reply", {
    body: "New text",
    replace_draft_id: "d1",
  });
  expect(w.bodyOf("PATCH /me/messages/d1")).toEqual({
    body: { contentType: "text", content: "New text" },
  });
  expect(answer).not.toContain("edited");
});

const EVERY: [string, Record<string, unknown>][] = [
  ["outlook_send", SEND],
  ["outlook_reply", { id: "m1", body: "Thanks" }],
  ["outlook_reply_all", { id: "m3", body: "Thanks all." }],
  ["outlook_forward", { id: "m1", to: ["dana@example.com"], body: "FYI" }],
  ["outlook_draft_reply", { id: "m1", body: "A draft" }],
  ["outlook_draft_reply", { body: "New text", replace_draft_id: "d1" }],
];

for (const [tool, args] of EVERY) {
  test(`${tool} ${JSON.stringify(args)}: Cancel on the editable card writes nothing`, async () => {
    const w = world({ edit: null });
    expect(await w.fake.callTool(tool, args)).toBe(CANCELLED);
    expect(w.chooses.length).toBe(1);
    expect(w.writes()).toEqual([]);
    expect(w.graph.sent).toEqual([]);
    expect(w.fake.confirms).toEqual([]);
  });

  test(`${tool} ${JSON.stringify(args)}: the editable card is preferred over "choose-mail"`, async () => {
    const w = world({ edit: edits("draft", {}) });
    let plainAsked = false;
    Object.assign(w.fake.api, {
      features: [...(w.fake.api.features ?? []), "choose-mail"],
      chooseMail: async () => {
        plainAsked = true;
        return "send";
      },
    });
    const answer = await w.fake.callTool(tool, args);
    expect(plainAsked).toBe(false);
    expect(w.chooses.length).toBe(1);
    expect(answer).toStartWith("The reader saved it in Drafts: ");
    expect(w.graph.sent).toEqual([]);
  });

  test(`${tool} ${JSON.stringify(args)}: an old prifly (no new features) asks as it always did`, async () => {
    const w = world({});
    await w.fake.callTool(tool, args);
    expect(w.chooses).toEqual([]);
    expect(w.fake.confirms.length).toBe(1);
  });
}

test("an answer that is not the reader's own draft or send is a Cancel", async () => {
  for (const odd of [
    { choice: "yes", subject: "x", text: "y" },
    { choice: "send", text: "y" },
  ]) {
    const w = world({ edit: odd as never });
    expect(await w.fake.callTool("outlook_send", SEND)).toBe(CANCELLED);
    expect(w.writes()).toEqual([]);
  }
});

test("the editable card needs both the feature and the function", async () => {
  const absent = world({});
  Object.assign(absent.fake.api, {
    features: [...(absent.fake.api.features ?? []), "choose-mail-edit"],
  });
  await absent.fake.callTool("outlook_send", SEND);
  expect(absent.fake.confirms.length).toBe(1);

  const unlisted = world({});
  Object.assign(unlisted.fake.api, { chooseMailEdit: async () => null });
  await unlisted.fake.callTool("outlook_send", SEND);
  expect(unlisted.fake.confirms.length).toBe(1);
});

test("the editable card is asked bound to prifly's api", async () => {
  const w = world({});
  let bound = false;
  const api = w.fake.api;
  Object.assign(api, {
    features: [...(api.features ?? []), "choose-mail-edit"],
    chooseMailEdit(this: unknown) {
      bound = this === api;
      return Promise.resolve(null);
    },
  });
  expect(await w.fake.callTool("outlook_send", SEND)).toBe(CANCELLED);
  expect(bound).toBe(true);
});

test("confirm-edit, outlook_send: the card holds the body, and the edited text is what is sent", async () => {
  const w = world({ confirmEdit: (content) => `${content} See you there.` });
  w.fake.answer.confirm = false;
  const answer = await w.fake.callTool("outlook_send", SEND);
  expect(w.fake.confirms.length).toBe(1);
  const spec = w.fake.confirms[0] as { content?: string; summary?: string };
  expect(spec.content).toBe("Noon?");
  expect(spec.summary).toContain("To: bob@example.com");
  expect(spec.summary).toContain("Subject: Lunch?");
  expect(w.writes()).toEqual(["POST /me/sendMail"]);
  expect(w.graph.sent[0]?.["body"]).toEqual({
    contentType: "Text",
    content: "Noon? See you there.",
  });
  expect(answer).toStartWith('Sent "Lunch?" to bob@example.com');
  expect(answer).toEndWith('Its subject now: "Lunch?". Its text now:\nNoon? See you there.');
});

test("confirm-edit, outlook_reply: the quoted-mail line is shown but not part of the text", async () => {
  const w = world({ confirmEdit: (content) => content.toUpperCase() });
  await w.fake.callTool("outlook_reply", { id: "m1", body: "Thanks" });
  const spec = w.fake.confirms[0] as { content?: string; summary?: string };
  expect(spec.content).toBe("Thanks");
  expect(spec.summary).toContain("(Below it: the quoted mail from");
  expect(w.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "THANKS" });
  expect(w.writes()).toEqual(["POST /me/messages/m1/createReply", "POST /me/messages/draft1/send"]);
});

test("confirm-edit, unedited: written and answered exactly as with the plain confirm card", async () => {
  const edited = world({ confirmEdit: (content) => content });
  const plain = world({});
  for (const [tool, args] of EVERY) {
    expect(await edited.fake.callTool(tool, args), tool).toBe(
      await plain.fake.callTool(tool, args),
    );
  }
  expect(edited.graph.calls).toEqual(plain.graph.calls);
});

test("confirm-edit, Cancel writes nothing", async () => {
  for (const [tool, args] of EVERY) {
    const w = world({ confirmEdit: null });
    expect(await w.fake.callTool(tool, args), tool).toBe(CANCELLED);
    expect(w.writes()).toEqual([]);
  }
});

test("confirm-edit, outlook_draft_reply: the edited text is the reply draft's comment", async () => {
  const w = world({ confirmEdit: () => "Edited reply" });
  const answer = await w.fake.callTool("outlook_draft_reply", { id: "m1", body: "A draft" });
  const spec = w.fake.confirms[0] as { content?: string; summary?: string };
  expect(spec.content).toBe("A draft");
  expect(spec.summary).toContain("In reply to:");
  expect(w.writes()).toEqual(["POST /me/messages/m1/createReply"]);
  expect(w.bodyOf("POST /me/messages/m1/createReply")).toEqual({ comment: "Edited reply" });
  expect(answer).toStartWith('Reply draft saved in Drafts: "RE: Contract for review"');
  expect(answer).toEndWith("Its text now:\nEdited reply");
});

test("confirm-edit is not used where a two-way card exists", async () => {
  const w = world({ choose: "draft", confirmEdit: () => "ignored" });
  const answer = await w.fake.callTool("outlook_send", SEND);
  expect(answer).toStartWith("The reader saved it in Drafts: ");
  expect(w.fake.confirms).toEqual([]);
});
