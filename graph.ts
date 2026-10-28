/**
 * The Microsoft Graph client: every call goes through `api.connections.fetch`,
 * so prifly holds the token and this code never sees one. A call answers a
 * `GraphResult`: the value, or a typed `GraphFailure` for the caller to turn
 * into text (auth.ts does) — nothing here throws for a Graph problem.
 *
 * Graph's rules the calls below keep to:
 * - `$search` takes no `$filter` or `$orderby`, so its hits are sorted here;
 * - a `conversationId` filter combined with `$orderby` is refused as an
 *   inefficient filter, so a conversation is sorted here too.
 *
 * The writes (send, save, reply, forward, move, PATCH) are plain functions
 * here; the tools ask the reader on a card before they call any of them.
 */

import type { ExtensionConnectionsApi } from "@prifly/extension-api";
import {
  type Attachment,
  type Json,
  type Mail,
  newestFirst,
  oldestFirst,
  record,
  text,
  toAttachment,
  toMail,
} from "./model";

export const GRAPH_ORIGIN = "https://graph.microsoft.com";
const BASE = `${GRAPH_ORIGIN}/v1.0`;

/** The longest a `Retry-After` is waited out. */
export const MAX_WAIT_MS = 30_000;

const BODY_PREFER = 'outlook.body-content-type="text"';

const LIST_SELECT =
  "id,conversationId,subject,from,toRecipients,receivedDateTime,isRead,isDraft,hasAttachments," +
  "bodyPreview,webLink,flag,categories";
const FULL_SELECT = `${LIST_SELECT},ccRecipients,bccRecipients,replyTo,body`;

export type Wait = (ms: number) => Promise<void>;

/** One Microsoft connection, ready to call Graph with. */
export type GraphContext = {
  connections: ExtensionConnectionsApi;
  connection: string;
  /** Sleeps out a `Retry-After`; tests give one that does not sleep. */
  wait?: Wait;
};

export type GraphFailure =
  | {
      kind: "needs-sign-in";
      account: string;
      /** The tenant's admin can approve prifly here; set only when the host passed a link. */
      adminConsentUrl?: string;
    }
  /** A 401 that got past the host: a claims challenge or a token Graph refused. */
  | { kind: "unauthorized" }
  | { kind: "forbidden"; message: string }
  | { kind: "gone" }
  | { kind: "throttled"; status: number; retryAfterSeconds: number | null }
  | { kind: "http"; status: number; message: string }
  /** The host refused the call outright (an unknown connection, a host the token must not go to). */
  | { kind: "refused"; message: string };

export type GraphResult<T> = { ok: true; value: T } | { ok: false; failure: GraphFailure };

function good<T>(value: T): GraphResult<T> {
  return { ok: true, value };
}

function bad(failure: GraphFailure): { ok: false; failure: GraphFailure } {
  return { ok: false, failure };
}

const sleep: Wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** `Retry-After` as seconds, or null when it is missing or not a number of seconds. */
function retryAfter(response: Response): number | null {
  const seconds = Number(response.headers.get("Retry-After"));
  return Number.isFinite(seconds) && seconds >= 0 && response.headers.has("Retry-After")
    ? seconds
    : null;
}

async function attempt(
  ctx: GraphContext,
  url: string,
  init: RequestInit,
): Promise<Response | GraphFailure> {
  try {
    // A fresh Request each time: a body can be sent only once, and a retry sends it again.
    const answer = await ctx.connections.fetch(ctx.connection, new Request(url, init));
    if ("needsSignIn" in answer) {
      const { account, adminConsentUrl } = answer.needsSignIn;
      return {
        kind: "needs-sign-in",
        account,
        ...(adminConsentUrl === undefined ? {} : { adminConsentUrl }),
      };
    }
    return answer;
  } catch (error) {
    return { kind: "refused", message: error instanceof Error ? error.message : String(error) };
  }
}

async function errorMessage(response: Response): Promise<string> {
  const body = await response.text().catch(() => "");
  try {
    const error = record(record(JSON.parse(body))["error"]);
    const message = [text(error["code"]), text(error["message"])].filter((s) => s !== "");
    if (message.length > 0) return message.join(": ");
  } catch {
    // Not JSON: the raw text is the best there is.
  }
  return body.slice(0, 300) || response.statusText;
}

async function classify(response: Response): Promise<GraphFailure> {
  const { status } = response;
  if (status === 401) return { kind: "unauthorized" };
  if (status === 429 || status === 503) {
    return { kind: "throttled", status, retryAfterSeconds: retryAfter(response) };
  }
  const message = await errorMessage(response);
  if (status === 410) return { kind: "gone" };
  if (status === 403) return { kind: "forbidden", message };
  return { kind: "http", status, message };
}

/** One request, with the one wait-out of a 429 or 503 the plan allows. */
export async function send(
  ctx: GraphContext,
  pathOrUrl: string,
  init: RequestInit = {},
): Promise<GraphResult<Response>> {
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : `${BASE}${pathOrUrl}`;
  const headers = new Headers(init.headers);
  // A caller's own preference rides along with the body type.
  const own = headers.get("Prefer");
  headers.set("Prefer", own === null ? BODY_PREFER : `${BODY_PREFER}, ${own}`);
  if (init.body !== undefined) headers.set("Content-Type", "application/json");
  const options = { ...init, headers };
  let answer = await attempt(ctx, url, options);
  // A 503 may have been processed already: only a 429, or a request that changes nothing, is repeated.
  const repeatable = (status: number) =>
    status === 429 || (status === 503 && (init.method ?? "GET") === "GET");
  if (answer instanceof Response && repeatable(answer.status)) {
    const seconds = retryAfter(answer) ?? 5;
    await (ctx.wait ?? sleep)(Math.min(seconds * 1000, MAX_WAIT_MS));
    answer = await attempt(ctx, url, options);
  }
  if (!(answer instanceof Response)) return bad(answer);
  return answer.ok ? good(answer) : bad(await classify(answer));
}

async function sendJson(
  ctx: GraphContext,
  pathOrUrl: string,
  init: RequestInit = {},
): Promise<GraphResult<Json>> {
  const sent = await send(ctx, pathOrUrl, init);
  if (!sent.ok) return sent;
  try {
    return good(record(await sent.value.json()));
  } catch {
    return bad({ kind: "http", status: sent.value.status, message: "The answer was not JSON." });
  }
}

/** A request whose answer has no body that matters (202 Accepted, 204 No Content). */
async function sendNoContent(
  ctx: GraphContext,
  path: string,
  body: unknown,
): Promise<GraphResult<null>> {
  const sent = await send(ctx, path, { method: "POST", body: JSON.stringify(body) });
  return sent.ok ? good(null) : sent;
}

/** `$`-parameters as a query string: values percent-encoded, names left as Graph writes them. */
function query(params: Record<string, string | number>): string {
  return Object.entries(params)
    .map(([name, value]) => `${name}=${encodeURIComponent(String(value))}`)
    .join("&");
}

const id = encodeURIComponent;

async function mailList(
  ctx: GraphContext,
  path: string,
  params: Record<string, string | number>,
  order: (mails: Mail[]) => Mail[],
  pages = 1,
): Promise<GraphResult<Mail[]>> {
  const mails: Mail[] = [];
  let next = `${path}?${query(params)}`;
  for (let page = 0; page < pages && next !== ""; page++) {
    const answer = await sendJson(ctx, next);
    if (!answer.ok) return answer;
    const rows = answer.value["value"];
    if (Array.isArray(rows)) mails.push(...rows.map(toMail));
    next = text(answer.value["@odata.nextLink"]);
  }
  return good(order(mails));
}

const INBOX = "/me/mailFolders/inbox/messages";

export type Profile = { displayName: string; address: string };

export async function me(ctx: GraphContext): Promise<GraphResult<Profile>> {
  const answer = await sendJson(
    ctx,
    `/me?${query({ $select: "displayName,mail,userPrincipalName" })}`,
  );
  if (!answer.ok) return answer;
  const r = answer.value;
  return good({
    displayName: text(r["displayName"]),
    address: text(r["mail"]) || text(r["userPrincipalName"]),
  });
}

/** How many inbox messages are unread, as Outlook counts them. */
export async function unreadCount(ctx: GraphContext): Promise<GraphResult<number>> {
  const answer = await sendJson(
    ctx,
    `/me/mailFolders/inbox?${query({ $select: "unreadItemCount" })}`,
  );
  if (!answer.ok) return answer;
  const count = answer.value["unreadItemCount"];
  return good(typeof count === "number" ? count : 0);
}

/** The newest inbox mail. */
export function inbox(ctx: GraphContext, top: number): Promise<GraphResult<Mail[]>> {
  return mailList(
    ctx,
    INBOX,
    { $orderby: "receivedDateTime desc", $top: top, $select: LIST_SELECT },
    (mails) => mails,
  );
}

/** A KQL search over every folder, newest first. */
export function search(ctx: GraphContext, kql: string, top: number): Promise<GraphResult<Mail[]>> {
  const quoted = `"${kql.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
  return mailList(
    ctx,
    "/me/messages",
    { $search: quoted, $top: top, $select: LIST_SELECT },
    newestFirst,
  );
}

/** Every message of a conversation, bodies included, oldest first. */
export function conversation(
  ctx: GraphContext,
  conversationId: string,
): Promise<GraphResult<Mail[]>> {
  const quoted = conversationId.replace(/'/g, "''");
  return mailList(
    ctx,
    "/me/messages",
    { $filter: `conversationId eq '${quoted}'`, $top: 100, $select: FULL_SELECT },
    oldestFirst,
    20,
  );
}

/** One message with its body, recipients and link. */
export async function message(ctx: GraphContext, messageId: string): Promise<GraphResult<Mail>> {
  const answer = await sendJson(
    ctx,
    `/me/messages/${id(messageId)}?${query({ $select: FULL_SELECT })}`,
  );
  return answer.ok ? good(toMail(answer.value)) : answer;
}

/** A message's attachments; a reference attachment also gets its link. */
export async function attachments(
  ctx: GraphContext,
  messageId: string,
): Promise<GraphResult<Attachment[]>> {
  const base = `/me/messages/${id(messageId)}/attachments`;
  const select = "id,name,contentType,size,isInline";
  const answer = await sendJson(ctx, `${base}?${query({ $select: select })}`);
  if (!answer.ok) return answer;
  const rows = answer.value["value"];
  const found = Array.isArray(rows) ? rows.map(toAttachment) : [];
  const withLinks: Attachment[] = [];
  for (const attachment of found) {
    if (attachment.kind !== "reference") {
      withLinks.push(attachment);
      continue;
    }
    const detail = await sendJson(ctx, `${base}/${id(attachment.id)}`);
    withLinks.push(
      detail.ok ? { ...attachment, link: toAttachment(detail.value).link } : attachment,
    );
  }
  return good(withLinks);
}

/** The raw bytes of a file attachment, or the `.eml` of a forwarded mail (an item attachment). */
export async function downloadAttachment(
  ctx: GraphContext,
  messageId: string,
  attachmentId: string,
): Promise<GraphResult<Uint8Array>> {
  const sent = await send(
    ctx,
    `/me/messages/${id(messageId)}/attachments/${id(attachmentId)}/$value`,
  );
  if (!sent.ok) return sent;
  return good(new Uint8Array(await sent.value.arrayBuffer()));
}

/** A mail folder, with the path of display names that leads to it ("Inbox/Clients"). */
export type Folder = { id: string; name: string; path: string; unread: number; total: number };

const FOLDER_SELECT = "id,displayName,childFolderCount,unreadItemCount,totalItemCount";
const MAX_FOLDERS = 300;

const count = (value: unknown) => (typeof value === "number" ? value : 0);

/** Where to list a folder level from, and the path of the folder above it. */
type FolderLevel = { path: string; url: string };

function levelUrl(path: string): string {
  return `${path}?${query({ $top: 100, $select: FOLDER_SELECT })}`;
}

/** One level's folders, and the levels below them that have children. */
async function folderLevel(
  ctx: GraphContext,
  parent: FolderLevel,
): Promise<GraphResult<{ found: Folder[]; below: FolderLevel[] }>> {
  const answer = await sendJson(ctx, parent.url);
  if (!answer.ok) return answer;
  const rows = answer.value["value"];
  const found: Folder[] = [];
  const below: FolderLevel[] = [];
  for (const row of Array.isArray(rows) ? rows.map(record) : []) {
    const name = text(row["displayName"]);
    const folder: Folder = {
      id: text(row["id"]),
      name,
      path: parent.path === "" ? name : `${parent.path}/${name}`,
      unread: count(row["unreadItemCount"]),
      total: count(row["totalItemCount"]),
    };
    found.push(folder);
    if (count(row["childFolderCount"]) > 0) {
      below.push({
        path: folder.path,
        url: levelUrl(`/me/mailFolders/${id(folder.id)}/childFolders`),
      });
    }
  }
  return good({ found, below });
}

/** Every mail folder, children under their parents, at most three levels and 300 folders deep. */
export async function folders(ctx: GraphContext): Promise<GraphResult<Folder[]>> {
  const out: Folder[] = [];
  let level: FolderLevel[] = [{ path: "", url: levelUrl("/me/mailFolders") }];
  for (let depth = 0; depth < 3 && level.length > 0 && out.length < MAX_FOLDERS; depth++) {
    const next: FolderLevel[] = [];
    for (const parent of level) {
      const listed = await folderLevel(ctx, parent);
      if (!listed.ok) return listed;
      out.push(...listed.value.found);
      next.push(...listed.value.below);
    }
    level = next;
  }
  return good(out.slice(0, MAX_FOLDERS));
}

/** A recipient as Graph writes one. */
function recipient(address: string): Json {
  return { emailAddress: { address } };
}

/** A file to attach, read from disk already. */
export type OutgoingFile = { name: string; contentType: string; bytes: Uint8Array };

function fileAttachment(file: OutgoingFile): Json {
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: file.name,
    contentType: file.contentType,
    contentBytes: Buffer.from(file.bytes).toString("base64"),
  };
}

export type NewMail = {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  body: string;
  files: OutgoingFile[];
};

const messageJson = (mail: NewMail): Json => ({
  subject: mail.subject,
  body: { contentType: "Text", content: mail.body },
  toRecipients: mail.to.map(recipient),
  ccRecipients: mail.cc.map(recipient),
  bccRecipients: mail.bcc.map(recipient),
  attachments: mail.files.map(fileAttachment),
});

/** Sends a new mail at once (`sendMail`), saved to Sent Items. */
export function sendMail(ctx: GraphContext, mail: NewMail): Promise<GraphResult<null>> {
  return sendNoContent(ctx, "/me/sendMail", { message: messageJson(mail), saveToSentItems: true });
}

/** Saves a new mail in Drafts (`POST /me/messages`); nothing is sent. */
export async function saveMail(ctx: GraphContext, mail: NewMail): Promise<GraphResult<Mail>> {
  const body = JSON.stringify(messageJson(mail));
  const answer = await sendJson(ctx, "/me/messages", { method: "POST", body });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

export type ReplyKind = "reply" | "replyAll" | "forward";

const CREATE_ACTION: Record<ReplyKind, string> = {
  reply: "createReply",
  replyAll: "createReplyAll",
  forward: "createForward",
};

/**
 * A threaded draft in Drafts, `comment` above the quoted mail: a reply, a
 * reply to all, or a forward to `to`. Nothing is sent until `sendDraft`.
 */
export async function createDraft(
  ctx: GraphContext,
  messageId: string,
  kind: ReplyKind,
  comment: string,
  to: string[] = [],
): Promise<GraphResult<Mail>> {
  const body = kind === "forward" ? { comment, toRecipients: to.map(recipient) } : { comment };
  const answer = await sendJson(ctx, `/me/messages/${id(messageId)}/${CREATE_ACTION[kind]}`, {
    method: "POST",
    body: JSON.stringify(body),
  });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

/** Adds a file (up to 3 MB) to a draft. */
export async function addAttachment(
  ctx: GraphContext,
  draftId: string,
  file: OutgoingFile,
): Promise<GraphResult<null>> {
  const sent = await send(ctx, `/me/messages/${id(draftId)}/attachments`, {
    method: "POST",
    body: JSON.stringify(fileAttachment(file)),
  });
  return sent.ok ? good(null) : sent;
}

/** Sends a draft. */
export function sendDraft(ctx: GraphContext, draftId: string): Promise<GraphResult<null>> {
  return sendNoContent(ctx, `/me/messages/${id(draftId)}/send`, {});
}

/** Moves a message to a folder (an id or a well-known name); the moved copy has a new id. */
export async function moveMessage(
  ctx: GraphContext,
  messageId: string,
  destinationId: string,
): Promise<GraphResult<Mail>> {
  const answer = await sendJson(ctx, `/me/messages/${id(messageId)}/move`, {
    method: "POST",
    body: JSON.stringify({ destinationId }),
  });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

/** Changes a message's writable properties: `isRead`, `categories`, `flag`, a draft's `body`. */
export async function patchMessage(
  ctx: GraphContext,
  messageId: string,
  patch: Json,
): Promise<GraphResult<Mail>> {
  const answer = await sendJson(ctx, `/me/messages/${id(messageId)}`, {
    method: "PATCH",
    body: JSON.stringify(patch),
  });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

/** Replaces a draft's body with plain `body`, and its subject when one is given. */
export function updateDraft(
  ctx: GraphContext,
  draftId: string,
  body: string,
  subject?: string,
): Promise<GraphResult<Mail>> {
  const text = { body: { contentType: "text", content: body } };
  return patchMessage(ctx, draftId, subject === undefined ? text : { subject, ...text });
}
