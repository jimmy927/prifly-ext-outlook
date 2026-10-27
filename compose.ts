/**
 * The tools that write mail: a new mail (`sendMail`, or `POST /me/messages`
 * for a draft), and a reply, a reply to all or a forward of a mail
 * (`createReply` / `createReplyAll` / `createForward`, files added to that
 * draft, then `send` when it is sent). Each shows the reader exactly who gets
 * what on a card before the first write. Where prifly has the two-way mail
 * card, the reader chooses there to save it in Drafts or send it, whatever
 * the tool is called; an older prifly's confirm card does what the tool's
 * name says. A reply is sent only when the draft Outlook made is addressed to
 * the people the card showed.
 */

import { readFile, stat } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import type { ExtensionApi } from "@prifly/extension-api";
import { failureText, type Scope } from "./auth";
import {
  ACCOUNT_PROPERTY,
  askMail,
  type CardMail,
  canChooseMail,
  choiceSummary,
  inDrafts,
  list,
  type MailCard,
  type MailChoice,
  type Opened,
  open,
  personName,
  str,
  type ToolDeps,
  type ToolSpec,
  told,
} from "./common";
import {
  addAttachment,
  createDraft,
  message,
  type NewMail,
  type OutgoingFile,
  type ReplyKind,
  saveMail,
  sendDraft,
  sendMail,
} from "./graph";
import { addresses, type Mail, type Person, personLabel } from "./model";

/** Graph takes a file in one request up to about 4 MB, base64 included: 3 MB of file. */
export const MAX_FILE_BYTES = 3 * 1024 * 1024;

const TYPES: Record<string, string> = {
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".md": "text/markdown",
  ".html": "text/html",
  ".json": "application/json",
  ".zip": "application/zip",
  ".eml": "message/rfc822",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

function megabytes(bytes: number): string {
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function kilobytes(bytes: number): string {
  return `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

/** The files to attach, read from disk; or why not, in which case nothing is sent. */
export async function readFiles(paths: string[]): Promise<OutgoingFile[] | string> {
  const files: OutgoingFile[] = [];
  let total = 0;
  for (const given of paths) {
    const path = resolve(given);
    const info = await stat(path).catch(() => null);
    if (info === null || !info.isFile()) return `Not done: there is no file at ${path}.`;
    total += info.size;
    if (info.size > MAX_FILE_BYTES || total > MAX_FILE_BYTES) {
      return (
        `Not done: ${basename(path)} is ${megabytes(info.size)}, and attachments can be at most ` +
        `${megabytes(MAX_FILE_BYTES)} in all here. Share a link to the file instead. Nothing was sent.`
      );
    }
    const bytes = new Uint8Array(await readFile(path));
    const contentType = TYPES[extname(path).toLowerCase()] ?? "application/octet-stream";
    files.push({ name: basename(path), contentType, bytes });
  }
  return files;
}

function fileLines(files: OutgoingFile[]): string[] {
  return files.length === 0
    ? []
    : [`Attachments: ${files.map((f) => `${f.name} (${kilobytes(f.bytes.length)})`).join(", ")}`];
}

const looksLikeAddress = (value: string) => /^[^\s@<>]+@[^\s@<>]+$/.test(value);

/** The files as the card draws them, bytes included. */
export function cardFiles(files: OutgoingFile[]): NonNullable<CardMail["attachments"]> {
  return files.map((file) => ({
    filename: file.name,
    mimeType: file.contentType,
    size: file.bytes.length,
    bytes: file.bytes,
  }));
}

/** Who a mail goes to, as the card's mail fields: To always, Cc and Bcc when there are any. */
function cardPeople(
  to: string[],
  cc: string[],
  bcc: string[],
): Pick<CardMail, "to" | "cc" | "bcc"> {
  return {
    to: to.join(", "),
    ...(cc.length > 0 ? { cc: cc.join(", ") } : {}),
    ...(bcc.length > 0 ? { bcc: bcc.join(", ") } : {}),
  };
}

/**
 * The scopes a mail tool opens with: `own` for what its name does, and both
 * Mail.ReadWrite and Mail.Send where the reader may choose either.
 */
export function mailScopes(api: ExtensionApi, own: Scope[]): Scope[] {
  return canChooseMail(api) ? ["Mail.ReadWrite", "Mail.Send"] : own;
}

/** Writes a new mail as the reader chose: saved in Drafts, or sent at once. */
async function writeNew(opened: Opened, mail: NewMail, choice: MailChoice): Promise<string> {
  const everyone = [...mail.to, ...mail.cc, ...mail.bcc];
  if (choice.kind === "draft") {
    const saved = await saveMail(opened.ctx, mail);
    if (!saved.ok) return failureText(saved.failure, opened.account);
    return told(choice, inDrafts(saved.value, everyone));
  }
  const sent = await sendMail(opened.ctx, mail);
  if (!sent.ok) return failureText(sent.failure, opened.account);
  return told(
    choice,
    `Sent "${mail.subject}" to ${everyone.join(", ")} from ${opened.account}. A copy is in Sent Items.`,
  );
}

async function sendNew(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const to = list(args["to"]);
  const cc = list(args["cc"]);
  const bcc = list(args["bcc"]);
  const subject = str(args["subject"]);
  const body = str(args["body"]);
  if (to.length === 0) return "Give at least one address in to.";
  const wrong = [...to, ...cc, ...bcc].filter((each) => !looksLikeAddress(each));
  if (wrong.length > 0) return `Not done: not an email address: ${wrong.join(", ")}.`;
  if (subject.trim() === "" && body.trim() === "") return "Give a subject or a body.";
  const files = await readFiles(list(args["attachments"]));
  if (typeof files === "string") return files;
  const opened = await open(api, args, mailScopes(api, ["Mail.Send"]), deps);
  if (typeof opened === "string") return opened;
  const lines = [
    `To: ${to.join(", ")}`,
    ...(cc.length > 0 ? [`Cc: ${cc.join(", ")}`] : []),
    ...(bcc.length > 0 ? [`Bcc: ${bcc.join(", ")}`] : []),
    `Subject: ${subject}`,
    ...fileLines(files),
    "",
    body,
  ];
  const choice = await askMail(
    api,
    {
      title: "Send an email",
      summary: `Send this mail to ${[...to, ...cc, ...bcc].join(", ")}.`,
      content: lines.join("\n"),
      details: `POST /me/sendMail as ${opened.account}`,
    },
    "send",
    {
      title: "Write an email",
      summary: choiceSummary(subject, [...to, ...cc, ...bcc], files.length),
      details: `Save as draft: POST /me/messages. Send now: POST /me/sendMail. As ${opened.account}`,
      mail: { ...cardPeople(to, cc, bcc), subject, text: body, attachments: cardFiles(files) },
    },
  );
  if (typeof choice === "string") return choice;
  return writeNew(opened, { to, cc, bcc, subject, body, files }, choice);
}

const notMe = (account: string) => (person: Person) =>
  person.address.toLowerCase() !== account.toLowerCase();

/** Who Outlook will address a reply to, worked out from the original as Outlook does. */
export function replyRecipients(
  original: Mail,
  kind: ReplyKind,
  account: string,
): { to: Person[]; cc: Person[] } {
  const sender = original.replyTo.length > 0 ? original.replyTo : [original.from];
  if (kind === "reply") return { to: sender, cc: [] };
  const to = [...sender, ...original.to].filter(notMe(account));
  return { to: unique(to), cc: unique(original.cc.filter(notMe(account))) };
}

function unique(people: Person[]): Person[] {
  const seen = new Set<string>();
  return people.filter((person) => {
    const key = person.address.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

const PREFIX: Record<ReplyKind, string> = { reply: "RE", replyAll: "RE", forward: "FW" };
const TITLE: Record<ReplyKind, string> = {
  reply: "Send a reply",
  replyAll: "Send a reply to all",
  forward: "Forward an email",
};

const SUMMARY: Record<ReplyKind, string> = {
  reply: "Send this reply to",
  replyAll: "Send this reply to",
  forward: "Forward this mail to",
};
const ACTION: Record<ReplyKind, string> = {
  reply: "createReply",
  replyAll: "createReplyAll",
  forward: "createForward",
};

function sameSet(a: string[], b: string[]): boolean {
  const left = new Set(a);
  const right = new Set(b);
  return left.size === right.size && [...left].every((each) => right.has(each));
}

/** A reply, reply to all or forward, worked out and ready to show and write. */
export type Answer = {
  kind: ReplyKind;
  original: Mail;
  body: string;
  files: OutgoingFile[];
  forwardTo: string[];
  /** Who it goes to, as the card shows them. */
  people: { to: Person[]; cc: Person[] };
  /** Their addresses, lower-cased, to check the draft Outlook makes against. */
  shown: string[];
  subject: string;
};

/** Works out who an answer goes to and its subject, or says why it cannot be written. */
export function planAnswer(
  original: Mail,
  given: { kind: ReplyKind; body: string; files: OutgoingFile[]; forwardTo: string[] },
  account: string,
): Answer | string {
  const { kind } = given;
  const people =
    kind === "forward"
      ? { to: given.forwardTo.map((address) => ({ name: "", address })), cc: [] }
      : replyRecipients(original, kind, account);
  const shown = addresses([...people.to, ...people.cc]);
  if (shown.length === 0) return "Not done: that mail has nobody to reply to.";
  const subject = `${PREFIX[kind]}: ${original.subject.replace(/^(re|fw|fwd|aw|sv|vs):\s*/i, "")}`;
  return { ...given, original, people, shown, subject };
}

/** The line under the text that says the original is quoted below it. */
const quotedNote = (original: Mail) =>
  `(Below it: the quoted mail from ${personLabel(original.from)}, "${original.subject}".)`;

/** The two-way card for an answer: a neutral title, the mail, and the calls either way makes. */
export function answerCard(answer: Answer, account: string): MailCard {
  const { kind, original, people, files } = answer;
  return {
    title: kind === "forward" ? "Forward an email" : "Reply",
    summary: choiceSummary(
      answer.subject,
      [...people.to, ...people.cc].map(personName),
      files.length,
    ),
    details: `Answers message id ${original.id}: POST /me/messages/${original.id}/${ACTION[kind]}, then for Send now POST /me/messages/<draft id>/send, as ${account}`,
    mail: {
      ...cardPeople(people.to.map(personLabel), people.cc.map(personLabel), []),
      subject: answer.subject,
      text: `${answer.body}\n\n${quotedNote(original)}`,
      attachments: cardFiles(files),
    },
  };
}

/** The answer for an answer left in Drafts: where it is, and what to check when Outlook addressed it otherwise. */
function savedAnswer(draft: Mail, actual: string[], shown: string[], choice: MailChoice): string {
  const differs = sameSet(actual, shown)
    ? ""
    : ` Outlook addressed it to ${actual.join(", ")}, not to the ${shown.join(", ")} on the card: check it before it is sent.`;
  return told(choice, `${inDrafts(draft, actual)}${differs}`);
}

/**
 * Makes the draft, checks its recipients against the card, adds the files,
 * and sends it when that is what the reader chose (or the tool's name says).
 */
export async function writeAnswer(
  opened: Opened,
  answer: Answer,
  choice: MailChoice,
): Promise<string> {
  const { ctx, account } = opened;
  const { original, kind, body, forwardTo, shown } = answer;
  const draft = await createDraft(ctx, original.id, kind, body, forwardTo);
  if (!draft.ok) return failureText(draft.failure, account);
  const left = `The draft is in Drafts (id ${draft.value.id}) to check and send from Outlook.`;
  const actual = addresses([...draft.value.to, ...draft.value.cc]);
  const sending = choice.kind === "send";
  if (sending && !sameSet(actual, shown)) {
    return `Not sent: Outlook addressed it to ${actual.join(", ")}, not to the ${shown.join(", ")} you confirmed. ${left}`;
  }
  for (const file of answer.files) {
    const added = await addAttachment(ctx, draft.value.id, file);
    if (!added.ok) {
      const why = `${file.name} could not be attached: ${failureText(added.failure, account)} ${left}`;
      return sending ? `Not sent: ${why}` : told(choice, `Saved in Drafts, but ${why}`);
    }
  }
  if (!sending) return savedAnswer(draft.value, actual, shown, choice);
  const sent = await sendDraft(ctx, draft.value.id);
  if (!sent.ok) return `Not sent: ${failureText(sent.failure, account)} ${left}`;
  return told(
    choice,
    `Sent "${draft.value.subject}" to ${actual.join(", ")} from ${account}. A copy is in Sent Items.`,
  );
}

async function answerText(
  api: ExtensionApi,
  args: Record<string, unknown>,
  deps: ToolDeps,
  kind: ReplyKind,
) {
  const id = str(args["id"]);
  const body = str(args["body"]);
  const forwardTo = kind === "forward" ? list(args["to"]) : [];
  if (id === "") return "Give the id of the message (outlook_search lists them).";
  if (kind !== "forward" && body.trim() === "") return "Give the text of the reply.";
  if (kind === "forward" && forwardTo.length === 0) return "Give at least one address in to.";
  const wrong = forwardTo.filter((each) => !looksLikeAddress(each));
  if (wrong.length > 0) return `Not done: not an email address: ${wrong.join(", ")}.`;
  const files = await readFiles(list(args["attachments"]));
  if (typeof files === "string") return files;
  const opened = await open(api, args, ["Mail.ReadWrite", "Mail.Send"], deps);
  if (typeof opened === "string") return opened;
  const original = await message(opened.ctx, id);
  if (!original.ok) return failureText(original.failure, opened.account);
  const answer = planAnswer(original.value, { kind, body, files, forwardTo }, opened.account);
  if (typeof answer === "string") return answer;
  const { people, shown, subject } = answer;
  const lines = [
    `To: ${people.to.map(personLabel).join(", ")}`,
    ...(people.cc.length > 0 ? [`Cc: ${people.cc.map(personLabel).join(", ")}`] : []),
    `Subject: ${subject}`,
    ...fileLines(files),
    "",
    body,
    "",
    quotedNote(original.value),
  ];
  const choice = await askMail(
    api,
    {
      title: TITLE[kind],
      summary: `${SUMMARY[kind]} ${shown.join(", ")}.`,
      content: lines.join("\n"),
      details: `Answers message id ${id}: POST /me/messages/${id}/${ACTION[kind]}, then POST /me/messages/<draft id>/send, as ${opened.account}`,
    },
    "send",
    answerCard(answer, opened.account),
  );
  if (typeof choice === "string") return choice;
  return writeAnswer(opened, answer, choice);
}

const ATTACHMENTS = {
  type: "array",
  items: { type: "string" },
  description: "Paths of local files to attach, at most 3 MB in all.",
};

const ADDRESSES = (what: string) => ({
  type: "array",
  items: { type: "string" },
  description: `${what}: email addresses.`,
});

function replySpec(name: string, kind: ReplyKind, description: string): ToolSpec {
  return {
    name,
    description,
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "The id of the message, from outlook_search." },
        ...(kind === "forward" ? { to: ADDRESSES("Who to forward it to") } : {}),
        body: {
          type: "string",
          description:
            kind === "forward" ? "A note above the forwarded mail." : "The reply's text, plain.",
        },
        attachments: ATTACHMENTS,
        account: ACCOUNT_PROPERTY,
      },
      required: kind === "forward" ? ["id", "to"] : ["id", "body"],
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => answerText(api, args, deps, kind),
  };
}

/** What every mail tool's description ends with: who decides, and that there is no other way. */
export const READER_DECIDES =
  "The reader sees the mail (recipients, subject, text, files) on a card and chooses there to save it in Drafts or send it now, whatever this tool is called; the answer says which they chose (an older prifly's card only confirms what the tool's name says). Never try another way to send mail.";

export const composeTools: ToolSpec[] = [
  {
    name: "outlook_send",
    description: `Proposes a new Outlook mail, with optional local files attached. ${READER_DECIDES}`,
    inputSchema: {
      type: "object",
      properties: {
        to: ADDRESSES("Recipients"),
        cc: ADDRESSES("Cc"),
        bcc: ADDRESSES("Bcc"),
        subject: { type: "string" },
        body: { type: "string", description: "The text, plain." },
        attachments: ATTACHMENTS,
        account: ACCOUNT_PROPERTY,
      },
      required: ["to", "subject", "body"],
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => sendNew(api, args, deps),
  },
  replySpec(
    "outlook_reply",
    "reply",
    `Proposes a reply to the sender of an Outlook mail, threaded, the original quoted below. ${READER_DECIDES}`,
  ),
  replySpec(
    "outlook_reply_all",
    "replyAll",
    `Proposes a reply to everyone on an Outlook mail (sender, To and Cc), threaded. ${READER_DECIDES}`,
  ),
  replySpec(
    "outlook_forward",
    "forward",
    `Proposes forwarding an Outlook mail, with its attachments, to new recipients. ${READER_DECIDES}`,
  ),
];
