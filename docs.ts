/**
 * `outlook_api_docs`: reads the Microsoft Graph v1.0 reference on Microsoft
 * Learn, so a session can look up an endpoint before it calls `outlook_api`.
 * A topic or a Graph path is turned into a page name (a small table of the
 * mail and calendar pages, then a guess from the path's shape, then Learn's
 * own search), and the page's permissions, HTTP request, query parameters,
 * request body and response sections come back as plain text, with the
 * first example request.
 */

import { str, type ToolDeps, type ToolSpec } from "./common";
import { record, text } from "./model";

export const LEARN = "https://learn.microsoft.com/en-us/graph/";
const VIEW = "?view=graph-rest-1.0";
const MAX_TEXT = 14_000;

/** Topics by a squashed name (lower case, letters and digits), to their page under /graph/. */
const TOPICS: Record<string, string> = {
  sendmail: "api/user-sendmail",
  send: "api/message-send",
  senddraft: "api/message-send",
  reply: "api/message-reply",
  replyall: "api/message-replyall",
  forward: "api/message-forward",
  createreply: "api/message-createreply",
  createreplyall: "api/message-createreplyall",
  createforward: "api/message-createforward",
  createdraft: "api/user-post-messages",
  move: "api/message-move",
  copy: "api/message-copy",
  messages: "api/user-list-messages",
  listmessages: "api/user-list-messages",
  getmessage: "api/message-get",
  updatemessage: "api/message-update",
  deletemessage: "api/message-delete",
  message: "api/resources/message",
  attachments: "api/message-list-attachments",
  addattachment: "api/message-post-attachments",
  uploadsession: "api/attachment-createuploadsession",
  largeattachment: "api/attachment-createuploadsession",
  createuploadsession: "api/attachment-createuploadsession",
  folders: "api/user-list-mailfolders",
  mailfolders: "api/user-list-mailfolders",
  mailfolder: "api/resources/mailfolder",
  createfolder: "api/user-post-mailfolders",
  childfolders: "api/mailfolder-list-childfolders",
  rules: "api/mailfolder-list-messagerules",
  messagerules: "api/mailfolder-list-messagerules",
  categories: "api/outlookuser-list-mastercategories",
  mastercategories: "api/outlookuser-list-mastercategories",
  mailboxsettings: "api/user-get-mailboxsettings",
  automaticreplies: "api/user-update-mailboxsettings",
  outofoffice: "api/user-update-mailboxsettings",
  events: "api/user-list-events",
  listevents: "api/user-list-events",
  event: "api/resources/event",
  createevent: "api/user-post-events",
  getevent: "api/event-get",
  updateevent: "api/event-update",
  deleteevent: "api/event-delete",
  calendarview: "api/user-list-calendarview",
  calendars: "api/user-list-calendars",
  accept: "api/event-accept",
  decline: "api/event-decline",
  tentativelyaccept: "api/event-tentativelyaccept",
  cancelevent: "api/event-cancel",
  findmeetingtimes: "api/user-findmeetingtimes",
  getschedule: "api/calendar-getschedule",
  freebusy: "api/calendar-getschedule",
  delta: "api/message-delta",
  search: "search-query-parameter",
  filter: "filter-query-parameter",
  queryparameters: "query-parameters",
  throttling: "throttling",
};

/** A collection segment of a Graph path, by its lower-case name, to the resource it holds. */
const SINGULAR: Record<string, string> = {
  me: "user",
  messages: "message",
  mailfolders: "mailfolder",
  childfolders: "mailfolder",
  attachments: "attachment",
  events: "event",
  instances: "event",
  calendars: "calendar",
  calendargroups: "calendargroup",
  contacts: "contact",
  contactfolders: "contactfolder",
  messagerules: "messagerule",
  mastercategories: "outlookcategory",
  outlook: "outlookuser",
};

/** Actions that follow a collection directly, not an id (`/messages/delta`). */
const COLLECTION_ACTIONS = new Set(["delta", "createuploadsession"]);

/** Navigation properties that read like an action but are a resource of their own. */
const LIST_OF_ME = new Set(["calendarview", "mailboxsettings"]);

const VERB: Record<string, string> = {
  GET: "get",
  PATCH: "update",
  PUT: "update",
  DELETE: "delete",
};

type Token = { kind: "collection" | "id" | "action"; name: string };

function tokens(path: string): Token[] {
  const segments = path
    .split("?")[0]
    ?.split("/")
    .filter(
      (s) =>
        s !== "" &&
        s !== "v1.0" &&
        s !== "beta" &&
        !s.startsWith("https:") &&
        !s.includes("graph.microsoft.com"),
    );
  const out: Token[] = [];
  for (const segment of segments ?? []) {
    const lower = segment.toLowerCase();
    const previous = out.at(-1);
    if (lower in SINGULAR) out.push({ kind: "collection", name: lower });
    else if (
      previous?.kind === "collection" &&
      previous.name !== "me" &&
      previous.name !== "outlook" &&
      !COLLECTION_ACTIONS.has(lower)
    ) {
      out.push({ kind: "id", name: segment });
    } else out.push({ kind: "action", name: lower });
  }
  return out;
}

/** A page name guessed from a Graph path's shape: `POST /me/messages/{id}/move` is message-move. */
export function pageForPath(method: string, path: string): string | null {
  const parts = tokens(path);
  const last = parts.at(-1);
  if (last === undefined) return null;
  const owner = (before: number) => {
    const found = parts
      .slice(0, before)
      .reverse()
      .find((p) => p.kind === "collection");
    return found === undefined ? "user" : (SINGULAR[found.name] ?? "user");
  };
  if (last.kind === "action") {
    const resource = owner(parts.length - 1);
    if (LIST_OF_ME.has(last.name)) {
      return `api/${resource}-${last.name === "calendarview" ? "list" : (VERB[method] ?? "get")}-${last.name}`;
    }
    return `api/${resource}-${last.name}`;
  }
  if (last.kind === "id") {
    const collection = parts.at(-2);
    return `api/${SINGULAR[collection?.name ?? ""] ?? "user"}-${VERB[method] ?? "get"}`;
  }
  if (last.name === "me") return "api/user-get";
  const verb = method === "POST" ? "post" : "list";
  return `api/${owner(parts.length - 1)}-${verb}-${last.name}`;
}

/** The page a topic names, before any search: a Learn URL, a page name, a path, or a known topic. */
export function pageFor(topic: string): string | null {
  const given = topic.trim();
  const learn = /learn\.microsoft\.com\/[a-z-]+\/graph\/([^?#]+)/i.exec(given);
  if (learn?.[1] !== undefined) return learn[1];
  if (/^(resources\/)?[a-z0-9]+(-[a-z0-9]+)+$/i.test(given)) return `api/${given.toLowerCase()}`;
  const request = /^(GET|POST|PATCH|PUT|DELETE)?\s*(\S*\/\S*)$/i.exec(given);
  if (request?.[2] !== undefined)
    return pageForPath((request[1] ?? "GET").toUpperCase(), request[2]);
  return TOPICS[given.toLowerCase().replace(/[^a-z0-9]/g, "")] ?? null;
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

function decode(html: string): string {
  return html.replace(/&(#x?[0-9a-f]+|[a-z]+);/gi, (whole, name: string) => {
    if (name.startsWith("#x") || name.startsWith("#X"))
      return String.fromCodePoint(Number.parseInt(name.slice(2), 16));
    if (name.startsWith("#")) return String.fromCodePoint(Number(name.slice(1)));
    return ENTITIES[name.toLowerCase()] ?? whole;
  });
}

/** A fragment of a Learn page as plain text: code blocks kept as they are, tables as `| a | b |` rows. */
export function htmlToText(html: string): string {
  // Code keeps its own spacing; everywhere else whitespace is insignificant, as in HTML.
  const code: string[] = [];
  const held = html.replace(
    /<pre[^>]*>\s*<code[^>]*>([\s\S]*?)<\/code>\s*<\/pre>/gi,
    (_, inner: string) => {
      code.push(decode(inner.replace(/<[^>]+>/g, "")).trimEnd());
      // A private-use character marks where the block goes back: it never appears on a Learn page.
      return `${code.length - 1}`;
    },
  );
  const plain = held
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/\s+/g, " ")
    .replace(/<h[1-6][^>]*>\s*/gi, "\n\n## ")
    .replace(/<tr[^>]*>\s*/gi, "\n| ")
    .replace(/\s*<\/t[dh]>\s*/gi, " | ")
    .replace(/<li[^>]*>\s*/gi, "\n- ")
    .replace(/<(br|\/p|\/div|\/h[1-6]|\/table|\/ul|\/ol)[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, "");
  return decode(plain)
    .split("\n")
    .map((line) => line.trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/(\d+)/g, (_, at: string) => `\n\`\`\`\n${code[Number(at)] ?? ""}\n\`\`\`\n`)
    .trim();
}

const SECTIONS = new Set([
  "permissions",
  "http-request",
  "function-parameters",
  "path-parameters",
  "optional-query-parameters",
  "query-parameters",
  "request-headers",
  "request-body",
  "response",
  "properties",
  "relationships",
  "methods",
  "json-representation",
]);

/** The reference sections of a Learn page, its first example request, or its main text. */
export function referenceText(html: string): string {
  const title = htmlToText(/<h1[^>]*>([\s\S]*?)<\/h1>/i.exec(html)?.[1] ?? "");
  const heads = [...html.matchAll(/<h2 id="([^"]+)"[^>]*>/gi)];
  const parts: string[] = [];
  heads.forEach((head, at) => {
    const id = head[1] ?? "";
    const end = heads[at + 1]?.index ?? html.length;
    const section = html.slice(head.index, end);
    if (SECTIONS.has(id)) parts.push(htmlToText(section));
    else if (id.startsWith("example") && !parts.some((p) => p.startsWith("## Example request"))) {
      const code = /<pre[^>]*>\s*<code class="lang-http"[^>]*>([\s\S]*?)<\/code>/i.exec(section);
      if (code?.[1] !== undefined)
        parts.push(`## Example request\n\`\`\`\n${decode(code[1]).trim()}\n\`\`\``);
    }
  });
  if (parts.length === 0) {
    // Not a method or resource page (a guide such as search-query-parameter): its article text,
    // from below the title to the feedback block, without the page's own chrome.
    const start = html.search(/<\/h1>/i);
    const end = html.search(/<h2 id="ms--feedback"|<\/main>/i);
    const article = html.slice(start < 0 ? 0 : start, end < 0 ? html.length : end);
    parts.push(htmlToText(article.replace(/<(nav|button|form|script|style)[\s\S]*?<\/\1>/gi, "")));
  }
  const body = parts.join("\n\n");
  const cut =
    body.length > MAX_TEXT ? `${body.slice(0, MAX_TEXT)}\n… cut here; the page has more.` : body;
  return title === "" ? cut : `# ${title}\n\n${cut}`;
}

type Fetch = (url: string) => Promise<Response>;

/** Learn's search, kept to Graph's reference: the first method page, else the first resource page. */
async function searchPage(fetchDocs: Fetch, topic: string): Promise<string | null> {
  const query = new URLSearchParams({
    search: `${topic} microsoft graph api`,
    locale: "en-us",
    $top: "25",
  });
  const answer = await fetchDocs(`https://learn.microsoft.com/api/search?${query.toString()}`);
  if (!answer.ok) return null;
  const results = record(await answer.json().catch(() => ({})))["results"];
  const pages = (Array.isArray(results) ? results : [])
    .map(
      (result) =>
        /learn\.microsoft\.com\/en-us\/graph\/(api\/[^?#]+)/.exec(text(record(result)["url"]))?.[1],
    )
    .filter((page): page is string => page !== undefined && page !== "api/overview");
  return pages.find((page) => !page.startsWith("api/resources/")) ?? pages[0] ?? null;
}

async function fetchPage(
  fetchDocs: Fetch,
  page: string,
): Promise<{ url: string; html: string } | null> {
  const url = `${LEARN}${page}${VIEW}`;
  const answer = await fetchDocs(url);
  return answer.ok ? { url, html: await answer.text() } : null;
}

async function docsText(args: Record<string, unknown>, deps: ToolDeps): Promise<string> {
  const topic = str(args["topic"]).trim();
  if (topic === "")
    return "Give a topic (sendMail, create event) or a Graph path (POST /me/messages/{id}/move).";
  const fetchDocs: Fetch = deps.fetchDocs ?? ((url) => fetch(url));
  const guessed = pageFor(topic);
  let found = guessed === null ? null : await fetchPage(fetchDocs, guessed);
  if (found === null) {
    const searched = await searchPage(fetchDocs, topic);
    found = searched === null || searched === guessed ? null : await fetchPage(fetchDocs, searched);
  }
  if (found === null) {
    return (
      `No Microsoft Graph v1.0 reference page found for "${topic}"${guessed === null ? "" : ` (tried ${guessed})`}. ` +
      `Give a page name such as message-move or user-list-events, or start from ${LEARN}api/overview${VIEW}.`
    );
  }
  return `From ${found.url}\n\n${referenceText(found.html)}`;
}

export const docsTool: ToolSpec = {
  name: "outlook_api_docs",
  description:
    "Read the Microsoft Graph v1.0 reference for an endpoint from Microsoft Learn: permissions, the HTTP request, " +
    "query parameters, the request body, the response and an example request. topic is a name (sendMail, " +
    "create event, calendarView), a Graph path (POST /me/messages/{id}/move), a page name (message-move) or a Learn URL.",
  inputSchema: {
    type: "object",
    properties: {
      topic: {
        type: "string",
        description: "What to look up: a name, a Graph path, a page name or a Learn URL.",
      },
    },
    required: ["topic"],
    additionalProperties: false,
  },
  run: (_api, args, _ctx, deps) => docsText(args, deps),
};
