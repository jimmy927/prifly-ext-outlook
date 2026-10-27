/**
 * The MCP tools every session sees as `mcp__prifly__outlook_*`. Reads run at
 * once; every write shows the reader a card first (`ask`, or for a mail
 * `askMail`, where the reader chooses draft or send) and makes no write call
 * when they cancel. This file holds the reads and the draft reply,
 * and puts every tool together; the writes live in compose.ts (send, reply,
 * forward) and changes.ts (move, flag, categories, read state, delete), the
 * generic Graph call in api-tool.ts and the reference reader in docs.ts.
 */

import type { ExtensionApi, ExtensionTool, ExtensionToolContext } from "@prifly/extension-api";
import { apiTools } from "./api-tool";
import { failureText } from "./auth";
import { changeTools } from "./changes";
import {
  ACCOUNT_PROPERTY,
  ask,
  askMail,
  type CardMail,
  canChooseMail,
  choiceSummary,
  clamp,
  DATA_NOTE,
  inDrafts,
  type MailCard,
  type Opened,
  open,
  personName,
  quoted,
  str,
  type ToolDeps,
  type ToolSpec,
  told,
  when,
} from "./common";
import {
  answerCard,
  composeTools,
  mailScopes,
  planAnswer,
  READER_DECIDES,
  writeAnswer,
} from "./compose";
import { docsTool } from "./docs";
import {
  attachments,
  conversation,
  createDraft,
  folders,
  inbox,
  message,
  type ReplyKind,
  search,
  sendDraft,
  updateDraft,
} from "./graph";
import { addresses, type Mail, oldestFirst, type Person, personLabel } from "./model";
import { attachmentDir, saveAttachments } from "./save-attachments";

export type { ToolDeps } from "./common";

async function searchText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const opened = await open(api, args, ["Mail.Read"], deps);
  if (typeof opened === "string") return opened;
  const query = str(args["query"]).trim();
  const max = clamp(args["max"], 15, 50);
  const found = query === "" ? await inbox(opened.ctx, max) : await search(opened.ctx, query, max);
  if (!found.ok) return failureText(found.failure, opened.account);
  if (found.value.length === 0) return "No mail matched.";
  const rows = found.value.map((mail) =>
    [
      mail.id,
      when(mail.received),
      personLabel(mail.from),
      mail.subject.replaceAll("|", "/"),
      mail.hasAttachments ? "attachments" : "-",
      mail.isRead ? "read" : "unread",
    ].join(" | "),
  );
  return [DATA_NOTE, "id | date | from | subject | attachments | read", ...rows].join("\n");
}

function header(mail: Mail): string[] {
  const lines = [`From: ${personLabel(mail.from)}`, `To: ${mail.to.map(personLabel).join(", ")}`];
  if (mail.cc.length > 0) lines.push(`Cc: ${mail.cc.map(personLabel).join(", ")}`);
  lines.push(`Date: ${when(mail.received)}`, `Subject: ${mail.subject}`, `Id: ${mail.id}`);
  if (mail.categories.length > 0) lines.push(`Categories: ${mail.categories.join(", ")}`);
  if (mail.flag.status !== "notFlagged") lines.push(`Flag: ${mail.flag.status}`);
  if (mail.webLink !== "") lines.push(`Open in Outlook: ${mail.webLink}`);
  return lines;
}

async function readText(
  api: ExtensionApi,
  args: Record<string, unknown>,
  ctx: ExtensionToolContext,
  deps: ToolDeps,
) {
  const id = str(args["id"]);
  if (id === "") return "Give the id of a message (outlook_search lists them).";
  const opened = await open(api, args, ["Mail.Read"], deps);
  if (typeof opened === "string") return opened;
  const first = await message(opened.ctx, id);
  if (!first.ok) return failureText(first.failure, opened.account);
  let mails = [first.value];
  if (args["conversation"] === true) {
    const all = await conversation(opened.ctx, first.value.conversationId);
    if (!all.ok) return failureText(all.failure, opened.account);
    mails = all.value.length > 0 ? oldestFirst(all.value) : mails;
  }
  const dir = attachmentDir(api, ctx, first.value.conversationId);
  const counter = { next: 1 };
  const parts = [DATA_NOTE];
  for (const mail of mails) {
    const files = await saveAttachments(opened.ctx, opened.account, dir, mail, counter);
    const attached = files.length > 0 ? ["", "Attachments:", ...files] : [];
    parts.push(["", "---", ...header(mail), "", mail.body.trim(), ...attached].join("\n"));
  }
  return parts.join("\n");
}

async function foldersText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const opened = await open(api, args, ["Mail.Read"], deps);
  if (typeof opened === "string") return opened;
  const found = await folders(opened.ctx);
  if (!found.ok) return failureText(found.failure, opened.account);
  const rows = found.value.map((f) => `${f.path} | ${f.unread} | ${f.total} | ${f.id}`);
  return ["folder | unread | total | id", ...rows].join("\n");
}

/** The files already on a draft, as the two-way card lists them; none when prifly has no such card. */
async function draftFiles(api: ExtensionApi, opened: Opened, draft: Mail) {
  if (!draft.hasAttachments || !canChooseMail(api)) return [];
  const found = await attachments(opened.ctx, draft.id);
  return found.ok
    ? found.value.map((each) => ({
        filename: each.name,
        mimeType: each.contentType,
        size: each.size,
      }))
    : [];
}

/** The two-way card for a draft whose text is replaced: the draft as it will be. */
function rewriteCard(draft: Mail, body: string, files: CardMail["attachments"]): MailCard {
  const label = (people: Person[]) => people.map(personLabel).join(", ");
  return {
    title: "Reply",
    summary: choiceSummary(
      draft.subject,
      [...draft.to, ...draft.cc, ...draft.bcc].map(personName),
      files?.length ?? 0,
    ),
    details: `PATCH /me/messages/${draft.id}, then for Send now POST /me/messages/${draft.id}/send`,
    mail: {
      to: label(draft.to),
      ...(draft.cc.length > 0 ? { cc: label(draft.cc) } : {}),
      ...(draft.bcc.length > 0 ? { bcc: label(draft.bcc) } : {}),
      subject: draft.subject,
      text: body,
      ...(files === undefined || files.length === 0 ? {} : { attachments: files }),
    },
  };
}

/** Refuses to overwrite anything that is not a draft: a received mail's body is not ours to change. */
async function replaceDraft(
  api: ExtensionApi,
  opened: Opened,
  draftId: string,
  body: string,
): Promise<string> {
  const current = await message(opened.ctx, draftId);
  if (!current.ok) return failureText(current.failure, opened.account);
  if (!current.value.isDraft) return "That id is not a draft; nothing was changed.";
  const choice = await askMail(
    api,
    {
      title: "Change a draft",
      summary: `Replace the text of the draft ${quoted(current.value.subject)} in Drafts.`,
      content: `New text:\n${body}`,
      details: `PATCH /me/messages/${draftId}`,
    },
    "draft",
    rewriteCard(current.value, body, await draftFiles(api, opened, current.value)),
  );
  if (typeof choice === "string") return choice;
  const updated = await updateDraft(opened.ctx, draftId, body);
  if (!updated.ok) return failureText(updated.failure, opened.account);
  const draft = updated.value;
  const everyone = addresses([...draft.to, ...draft.cc, ...draft.bcc]);
  if (choice.kind === "draft") {
    return choice.chosen
      ? told(choice, inDrafts(draft, everyone))
      : draftSaved("Draft updated", draft);
  }
  const sent = await sendDraft(opened.ctx, draftId);
  if (!sent.ok) {
    return `Not sent: ${failureText(sent.failure, opened.account)} The draft is in Drafts (id ${draftId}) with the new text.`;
  }
  return told(
    choice,
    `Sent "${draft.subject}" to ${everyone.join(", ")} from ${opened.account}. A copy is in Sent Items.`,
  );
}

function draftSaved(what: string, draft: Mail): string {
  const link = draft.webLink === "" ? "" : `\nOpen it in Outlook: ${draft.webLink}`;
  return (
    `${what} in Drafts: "${draft.subject}". Draft id: ${draft.id}${link}\n` +
    "Nothing was sent: the reader sends it from Outlook, or outlook_reply sends a reply directly."
  );
}

async function draftText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const id = str(args["id"]);
  const body = str(args["body"]);
  const replace = str(args["replace_draft_id"]);
  if (body.trim() === "" || (id === "" && replace === "")) {
    return "Give the id of the message to answer and the text of the reply.";
  }
  const opened = await open(api, args, mailScopes(api, ["Mail.ReadWrite"]), deps);
  if (typeof opened === "string") return opened;
  if (replace !== "") return replaceDraft(api, opened, replace, body);
  const original = await message(opened.ctx, id);
  if (!original.ok) return failureText(original.failure, opened.account);
  const all = args["all"] === true;
  if (canChooseMail(api)) return chosenReply(api, opened, original.value, all, body);
  const refused = await ask(api, {
    title: "Save a reply draft",
    summary: `Save a ${all ? "reply-all" : "reply"} to ${quoted(original.value.subject)} in Drafts, not sent.`,
    content: `In reply to: ${personLabel(original.value.from)}, ${quoted(original.value.subject)}\n\n${body}`,
    details: `POST /me/messages/${id}/${all ? "createReplyAll" : "createReply"}`,
  });
  if (refused !== null) return refused;
  const created = await createDraft(opened.ctx, id, all ? "replyAll" : "reply", body);
  if (!created.ok) return failureText(created.failure, opened.account);
  return draftSaved("Reply draft saved", created.value);
}

/**
 * A reply draft on prifly's two-way card: the reader chooses Save as draft
 * or Send now, and it is written as outlook_reply writes it, sent only on
 * their Send.
 */
async function chosenReply(
  api: ExtensionApi,
  opened: Opened,
  original: Mail,
  all: boolean,
  body: string,
): Promise<string> {
  const kind: ReplyKind = all ? "replyAll" : "reply";
  const answer = planAnswer(original, { kind, body, files: [], forwardTo: [] }, opened.account);
  if (typeof answer === "string") return answer;
  const card = answerCard(answer, opened.account);
  // Without "choose-mail" this spec would be a confirm card; askMail is only reached with it.
  const choice = await askMail(
    api,
    { title: card.title, summary: card.summary, details: card.details },
    "draft",
    card,
  );
  if (typeof choice === "string") return choice;
  return writeAnswer(opened, answer, choice);
}

const readTools: ToolSpec[] = [
  {
    name: "outlook_search",
    description:
      "Search the reader's Outlook mail with KQL (from:ada subject:invoice received>=2026-10-01 hasattachments:true isread:false). " +
      "An empty query lists the newest inbox mail. Answers id | date | from | subject | attachments | read, newest first.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "A KQL search; empty for the newest inbox mail." },
        max: { type: "number", description: "How many results, 1 to 50 (default 15)." },
        account: ACCOUNT_PROPERTY,
      },
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => searchText(api, args, deps),
  },
  {
    name: "outlook_read",
    description:
      "Read one Outlook mail by id (from outlook_search), or with conversation true the whole conversation, oldest first. " +
      "Attachments are saved to disk and their paths listed.",
    inputSchema: {
      type: "object",
      properties: {
        id: { type: "string", description: "The message id outlook_search gave." },
        conversation: { type: "boolean", description: "Read every message of its conversation." },
        account: ACCOUNT_PROPERTY,
      },
      required: ["id"],
      additionalProperties: false,
    },
    run: readText,
  },
  {
    name: "outlook_folders",
    description:
      "List the reader's mail folders (three levels deep) as folder path | unread | total | id. " +
      "outlook_move takes a folder's id, its path or name, or a well-known name such as archive.",
    inputSchema: {
      type: "object",
      properties: { account: ACCOUNT_PROPERTY },
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => foldersText(api, args, deps),
  },
  {
    name: "outlook_draft_reply",
    description:
      "Proposes a reply to an Outlook mail, threaded under it, to keep in Drafts. To change a draft made earlier, pass its id as replace_draft_id. " +
      READER_DECIDES,
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "The id of the message to answer; ignored when replace_draft_id is given.",
        },
        body: { type: "string", description: "The reply's text, plain." },
        all: {
          type: "boolean",
          description: "Reply to everyone on the mail (default: the sender).",
        },
        replace_draft_id: {
          type: "string",
          description: "The id of a draft: its body is replaced instead of a new draft.",
        },
        account: ACCOUNT_PROPERTY,
      },
      required: ["body"],
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => draftText(api, args, deps),
  },
];

/** Runs one tool's work, and turns anything it throws into text. */
async function guarded(work: () => Promise<string>): Promise<string> {
  try {
    return await work();
  } catch (error) {
    return `Outlook could not do that: ${error instanceof Error ? error.message : String(error)}`;
  }
}

export function outlookTools(
  getApi: () => ExtensionApi | null,
  deps: ToolDeps = {},
): ExtensionTool[] {
  const specs = [...readTools, ...composeTools, ...changeTools, ...apiTools, docsTool];
  return specs.map((spec) => ({
    name: spec.name,
    description: spec.description,
    inputSchema: spec.inputSchema,
    call: (args, ctx) =>
      guarded(async () => {
        const api = getApi();
        return api === null
          ? "The Outlook extension is not running."
          : spec.run(api, args, ctx, deps);
      }),
  }));
}
