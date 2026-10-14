/**
 * The tools that send mail: a new mail (`sendMail`), and a reply, a reply to
 * all or a forward of a mail (`createReply` / `createReplyAll` /
 * `createForward`, files added to that draft, then `send`). Each shows the
 * reader exactly who gets what on a confirm card before the first write,
 * and a reply is sent only when the draft Outlook made is addressed to the
 * people the card showed.
 */

import { readFile, stat } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import type { ExtensionApi } from "@prifly/extension-api";
import { failureText } from "./auth";
import {
  ACCOUNT_PROPERTY,
  ask,
  list,
  type Opened,
  open,
  str,
  type ToolDeps,
  type ToolSpec,
} from "./common";
import {
  addAttachment,
  addresses,
  createDraft,
  message,
  type OutgoingFile,
  type ReplyKind,
  sendDraft,
  sendMail,
} from "./graph";
import { type Mail, type Person, personLabel } from "./model";

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
  const opened = await open(api, args, ["Mail.Send"], deps);
  if (typeof opened === "string") return opened;
  const lines = [
    `From: ${opened.account}`,
    `To: ${to.join(", ")}`,
    ...(cc.length > 0 ? [`Cc: ${cc.join(", ")}`] : []),
    ...(bcc.length > 0 ? [`Bcc: ${bcc.join(", ")}`] : []),
    `Subject: ${subject}`,
    ...fileLines(files),
    "",
    body,
  ];
  const refused = await ask(api, {
    title: "Send an email",
    summary: `Send "${subject}" to ${[...to, ...cc, ...bcc].join(", ")}`,
    details: lines.join("\n"),
  });
  if (refused !== null) return refused;
  const sent = await sendMail(opened.ctx, { to, cc, bcc, subject, body, files });
  if (!sent.ok) return failureText(sent.failure, opened.account);
  return `Sent "${subject}" to ${[...to, ...cc, ...bcc].join(", ")} from ${opened.account}. A copy is in Sent Items.`;
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

function sameSet(a: string[], b: string[]): boolean {
  const left = new Set(a);
  const right = new Set(b);
  return left.size === right.size && [...left].every((each) => right.has(each));
}

type Answer = { kind: ReplyKind; body: string; files: OutgoingFile[]; forwardTo: string[] };

/** Makes the draft, checks its recipients against the card, adds the files and sends it. */
async function sendAnswer(
  opened: Opened,
  original: Mail,
  answer: Answer,
  shown: string[],
): Promise<string> {
  const { ctx, account } = opened;
  const draft = await createDraft(ctx, original.id, answer.kind, answer.body, answer.forwardTo);
  if (!draft.ok) return failureText(draft.failure, account);
  const left = `The draft is in Drafts (id ${draft.value.id}) to check and send from Outlook.`;
  const actual = addresses([...draft.value.to, ...draft.value.cc]);
  if (!sameSet(actual, shown)) {
    return `Not sent: Outlook addressed it to ${actual.join(", ")}, not to the ${shown.join(", ")} you confirmed. ${left}`;
  }
  for (const file of answer.files) {
    const added = await addAttachment(ctx, draft.value.id, file);
    if (!added.ok)
      return `Not sent: ${file.name} could not be attached: ${failureText(added.failure, account)} ${left}`;
  }
  const sent = await sendDraft(ctx, draft.value.id);
  if (!sent.ok) return `Not sent: ${failureText(sent.failure, account)} ${left}`;
  return `Sent "${draft.value.subject}" to ${actual.join(", ")} from ${account}. A copy is in Sent Items.`;
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
  const people =
    kind === "forward"
      ? { to: forwardTo.map((address) => ({ name: "", address })), cc: [] }
      : replyRecipients(original.value, kind, opened.account);
  const shown = addresses([...people.to, ...people.cc]);
  if (shown.length === 0) return "Not done: that mail has nobody to reply to.";
  const subject = `${PREFIX[kind]}: ${original.value.subject.replace(/^(re|fw|fwd|aw|sv|vs):\s*/i, "")}`;
  const lines = [
    `From: ${opened.account}`,
    `To: ${people.to.map(personLabel).join(", ")}`,
    ...(people.cc.length > 0 ? [`Cc: ${people.cc.map(personLabel).join(", ")}`] : []),
    `Subject: ${subject}`,
    ...fileLines(files),
    "",
    body,
    "",
    `(Below it: the quoted mail from ${personLabel(original.value.from)}, "${original.value.subject}".)`,
  ];
  const refused = await ask(api, {
    title: TITLE[kind],
    summary: `Send "${subject}" to ${shown.join(", ")}`,
    details: lines.join("\n"),
  });
  if (refused !== null) return refused;
  return sendAnswer(opened, original.value, { kind, body, files, forwardTo }, shown);
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

const CONFIRMS = "The reader confirms on a card that shows the recipients and the text first.";

export const composeTools: ToolSpec[] = [
  {
    name: "outlook_send",
    description: `Send a new Outlook mail, with optional local files attached. ${CONFIRMS}`,
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
    `Send a reply to the sender of an Outlook mail, threaded, the original quoted below. ${CONFIRMS}`,
  ),
  replySpec(
    "outlook_reply_all",
    "replyAll",
    `Send a reply to everyone on an Outlook mail (sender, To and Cc), threaded. ${CONFIRMS}`,
  ),
  replySpec(
    "outlook_forward",
    "forward",
    `Forward an Outlook mail, with its attachments, to new recipients. ${CONFIRMS}`,
  ),
];
