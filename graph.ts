/**
 * The Microsoft Graph client: every call goes through `api.connections.fetch`,
 * so prifly holds the token and this code never sees one. A call answers a
 * `GraphResult`: the value, or a typed `GraphFailure` for the caller to turn
 * into text (auth.ts does) — nothing here throws for a Graph problem.
 *
 * Graph's rules the calls below keep to:
 * - `$orderby`'s property must lead `$filter`, hence the dummy date in `unread`;
 * - `$search` takes no `$filter` or `$orderby`, so its hits are sorted here;
 * - a `conversationId` filter combined with `$orderby` is refused as an
 *   inefficient filter, so a conversation is sorted here too.
 *
 * The only writes are `createReply`, `createReplyAll` and a PATCH of a
 * draft's body. There is no call that sends.
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

const BASE = "https://graph.microsoft.com/v1.0";

/** The longest a `Retry-After` is waited out. */
export const MAX_WAIT_MS = 30_000;

const LIST_SELECT =
  "id,conversationId,subject,from,toRecipients,receivedDateTime,isRead,isDraft,hasAttachments," +
  "bodyPreview,webLink,flag,categories";
const FULL_SELECT = `${LIST_SELECT},ccRecipients,body`;

export type Wait = (ms: number) => Promise<void>;

/** One Microsoft connection, ready to call Graph with. */
export type GraphContext = {
  connections: ExtensionConnectionsApi;
  connection: string;
  /** Sleeps out a `Retry-After`; tests give one that does not sleep. */
  wait?: Wait;
};

export type GraphFailure =
  | { kind: "needs-sign-in"; account: string }
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
      return { kind: "needs-sign-in", account: answer.needsSignIn.account };
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
  if (status === 410 || /syncStateNotFound/i.test(message)) return { kind: "gone" };
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
  headers.set("Prefer", 'outlook.body-content-type="text"');
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

/** Unread inbox mail, newest first. */
export function unread(ctx: GraphContext, top: number): Promise<GraphResult<Mail[]>> {
  return mailList(
    ctx,
    INBOX,
    {
      $filter: "receivedDateTime ge 1900-01-01 and isRead eq false",
      $orderby: "receivedDateTime desc",
      $top: top,
      $select: LIST_SELECT,
    },
    (mails) => mails,
  );
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

/** Flagged mail from every folder, newest first; each carries its flag's due time. */
export function flagged(ctx: GraphContext, top: number): Promise<GraphResult<Mail[]>> {
  return mailList(
    ctx,
    "/me/messages",
    { $filter: "flag/flagStatus eq 'flagged'", $top: top, $select: LIST_SELECT },
    newestFirst,
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

/**
 * A threaded draft in Drafts, `comment` above the quoted mail. This creates a
 * draft only: Graph sends nothing until the reader sends it from Outlook.
 */
export async function createReply(
  ctx: GraphContext,
  messageId: string,
  all: boolean,
  comment: string,
): Promise<GraphResult<Mail>> {
  const action = all ? "createReplyAll" : "createReply";
  const answer = await sendJson(ctx, `/me/messages/${id(messageId)}/${action}`, {
    method: "POST",
    body: JSON.stringify({ comment }),
  });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

/** Replaces a draft's body with plain `body`. */
export async function updateDraft(
  ctx: GraphContext,
  draftId: string,
  body: string,
): Promise<GraphResult<Mail>> {
  const answer = await sendJson(ctx, `/me/messages/${id(draftId)}`, {
    method: "PATCH",
    body: JSON.stringify({ body: { contentType: "text", content: body } }),
  });
  return answer.ok ? good(toMail(answer.value)) : answer;
}

export type InboxDelta = {
  changed: Mail[];
  /** Ids of mail removed from the inbox since the last delta. */
  removed: string[];
  /** Pass it to the next `inboxDelta`; not a secret. */
  deltaLink: string;
};

const MAX_DELTA_PAGES = 50;

/**
 * The inbox's changes since `link` (the `deltaLink` of the last run), or all of
 * it when `link` is left out. A `gone` failure means the link expired: start again without one.
 */
export async function inboxDelta(
  ctx: GraphContext,
  link?: string,
): Promise<GraphResult<InboxDelta>> {
  let next = link ?? `${INBOX}/delta?${query({ $select: LIST_SELECT })}`;
  const changed: Mail[] = [];
  const removed: string[] = [];
  for (let page = 0; page < MAX_DELTA_PAGES; page++) {
    const answer = await sendJson(ctx, next);
    if (!answer.ok) return answer;
    const rows = answer.value["value"];
    for (const row of Array.isArray(rows) ? rows : []) {
      if ("@removed" in record(row)) removed.push(text(record(row)["id"]));
      else changed.push(toMail(row));
    }
    const deltaLink = text(answer.value["@odata.deltaLink"]);
    if (deltaLink !== "") return good({ changed, removed, deltaLink });
    next = text(answer.value["@odata.nextLink"]);
    if (next === "") break;
  }
  return bad({ kind: "http", status: 200, message: "The inbox delta never reached its end." });
}
