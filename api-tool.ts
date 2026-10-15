/**
 * `outlook_api`: any Microsoft Graph call with the reader's Microsoft
 * connection, for what the curated tools do not cover (calendar events,
 * rules, folders, mailbox settings…). A GET runs at once; any other method
 * shows the exact request on a confirm card first. Only
 * https://graph.microsoft.com is reached: another host is refused before
 * anything is sent.
 */

import type { ConfirmSpec, ExtensionApi } from "@prifly/extension-api";
import { failureText, type Scope } from "./auth";
import {
  ACCOUNT_PROPERTY,
  ask,
  DATA_NOTE,
  open,
  str,
  type ToolDeps,
  type ToolSpec,
} from "./common";
import { GRAPH_ORIGIN, send } from "./graph";
import { record } from "./model";

const METHODS = ["GET", "POST", "PATCH", "PUT", "DELETE"] as const;
type Method = (typeof METHODS)[number];

/** The longest answer handed back; Graph pages are cut here, with a note. */
const MAX_ANSWER = 60_000;

/** A full URL, kept only when it is on Graph's own origin with no user info. */
function absoluteUrl(given: string): URL | string {
  let url: URL;
  try {
    url = new URL(given);
  } catch {
    return `Not done: ${given} is not a URL.`;
  }
  if (url.origin !== GRAPH_ORIGIN || url.username !== "" || url.password !== "") {
    return `Not done: only ${GRAPH_ORIGIN} is reached, not ${url.origin}.`;
  }
  return url;
}

/** A path, made a v1.0 URL, kept only when it stays under /v1.0/me once `..` is resolved. */
function relativeUrl(given: string): URL | string {
  const relative = given.startsWith("/") ? given : `/${given}`;
  const versioned = /^\/v1\.0(\/|$)/.test(relative) ? relative : `/v1.0${relative}`;
  const url = new URL(`${GRAPH_ORIGIN}${versioned}`);
  if (
    url.origin === GRAPH_ORIGIN &&
    (url.pathname === "/v1.0/me" || url.pathname.startsWith("/v1.0/me/"))
  ) {
    return url;
  }
  return `Not done: a path must be under /v1.0/me, not ${url.pathname}; give a full ${GRAPH_ORIGIN} URL for anything else.`;
}

/** Turns `path` (a path under /me, or a full Graph URL) and `query` into the URL, or says why not. */
export function graphUrl(path: string, query: unknown): URL | string {
  const given = path.trim();
  if (given === "") {
    return "Give a path under /me (e.g. /me/messages) or a https://graph.microsoft.com URL.";
  }
  const url = /^[a-z][a-z0-9+.-]*:/i.test(given) ? absoluteUrl(given) : relativeUrl(given);
  if (typeof url === "string") return url;
  for (const [name, value] of Object.entries(record(query))) {
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      url.searchParams.set(name, String(value));
    }
  }
  return url;
}

/** The scopes a call needs, from its method and path. */
export function scopesFor(method: Method, url: URL): Scope[] {
  const path = url.pathname.toLowerCase();
  if (
    /\/(calendars?|events|calendarview|calendargroups|findmeetingtimes|getschedule)(\/|$)/.test(
      path,
    )
  ) {
    return ["Calendars.ReadWrite"];
  }
  if (/\/(sendmail|send|reply|replyall|forward)$/.test(path))
    return ["Mail.ReadWrite", "Mail.Send"];
  return method === "GET" ? ["Mail.Read"] : ["Mail.ReadWrite"];
}

const THINGS: Record<string, string> = {
  messages: "a message",
  events: "an event",
  mailfolders: "a mail folder",
  childfolders: "a mail folder",
  contacts: "a contact",
  contactfolders: "a contact folder",
  calendars: "a calendar",
  messagerules: "an inbox rule",
  attachments: "an attachment",
  categories: "a category",
};

const VERBS: Partial<Record<Method, string>> = {
  POST: "Create",
  PATCH: "Change",
  PUT: "Change",
  DELETE: "Delete",
};

/**
 * What a plain create, change or delete call does, in words ("Delete a
 * message"); null for anything else (an action such as /send, a path this does
 * not know), which the card then shows as the request body.
 */
export function readableLine(method: Method, url: URL): string | null {
  const parts = url.pathname.split("/").filter((part) => part !== "");
  const segments = parts.slice(parts[0]?.toLowerCase() === "v1.0" ? 1 : 0);
  if (segments[0]?.toLowerCase() === "me") segments.shift();
  const item = segments.length % 2 === 0;
  if (method === "POST" ? item : !item) return null;
  const thing = THINGS[(segments.at(item ? -2 : -1) ?? "").toLowerCase()];
  const verb = VERBS[method];
  return thing === undefined || verb === undefined ? null : `${verb} ${thing}`;
}

/** The confirm card for a write: the sentence and body for the reader, the request itself folded. */
function requestCard(
  method: Method,
  url: URL,
  body: string | undefined,
  account: string,
): ConfirmSpec {
  const line = readableLine(method, url);
  return {
    title: "Microsoft Graph request",
    summary: line === null ? `Send a ${method} request to Microsoft Graph.` : `${line}.`,
    content: line === null ? (body ?? "No request body.") : `${line} in the mailbox of ${account}.`,
    details: `${method} ${url.toString()}\n\n${body ?? "(no body)"}`,
  };
}

/** The body as it is sent: an object or array as JSON, a string as it is. */
function bodyText(body: unknown): string | undefined {
  if (body === undefined || body === null) return undefined;
  return typeof body === "string" ? body : JSON.stringify(body, null, 2);
}

/** A Graph answer as text: pretty JSON when it is JSON, cut at `MAX_ANSWER`. */
async function answerText(response: Response): Promise<string> {
  const raw = await response.text();
  let shown = raw;
  try {
    shown = raw === "" ? "" : JSON.stringify(JSON.parse(raw), null, 2);
  } catch {
    // Not JSON: as it came.
  }
  return shown.length > MAX_ANSWER
    ? `${shown.slice(0, MAX_ANSWER)}\n… cut at ${MAX_ANSWER} characters: ask for less ($select, $top).`
    : shown;
}

async function apiText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const method = (str(args["method"]) || "GET").toUpperCase() as Method;
  if (!METHODS.includes(method)) return `Not done: method is one of ${METHODS.join(", ")}.`;
  const url = graphUrl(str(args["path"]), args["query"]);
  if (typeof url === "string") return url;
  const body = bodyText(args["body"]);
  if (method === "GET" && body !== undefined) return "Not done: a GET takes no body.";
  const opened = await open(api, args, scopesFor(method, url), deps);
  if (typeof opened === "string") return opened;
  if (method !== "GET") {
    const refused = await ask(api, requestCard(method, url, body, opened.account));
    if (refused !== null) return refused;
  }
  const sent = await send(opened.ctx, url.toString(), {
    method,
    ...(body === undefined ? {} : { body }),
  });
  if (!sent.ok) return failureText(sent.failure, opened.account);
  const text = await answerText(sent.value);
  const head = `${method} ${url.toString()} answered ${sent.value.status}.`;
  return [head, ...(text === "" ? [] : [DATA_NOTE, text])].join("\n");
}

export const apiTools: ToolSpec[] = [
  {
    name: "outlook_api",
    description:
      "Call any Microsoft Graph v1.0 endpoint with the reader's Microsoft account: mail, folders, rules, mailbox settings, " +
      "calendars and events (/me/events, /me/calendarView). path is under /me (/me/messages?$top=5) or a full " +
      "https://graph.microsoft.com URL; no other host. A GET runs at once; POST, PATCH, PUT and DELETE show the reader " +
      "the request on a confirm card first. Look an endpoint up with outlook_api_docs before a write you have not made before.",
    inputSchema: {
      type: "object",
      properties: {
        method: { type: "string", enum: [...METHODS], description: "Default GET." },
        path: {
          type: "string",
          description:
            "/me/... (v1.0 is added), /v1.0/me/..., or a full https://graph.microsoft.com URL.",
        },
        query: {
          type: "object",
          description: 'Query parameters, e.g. {"$top": 10, "$select": "subject,from"}.',
        },
        body: { description: "The JSON body of a POST, PATCH or PUT." },
        account: ACCOUNT_PROPERTY,
      },
      required: ["path"],
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => apiText(api, args, deps),
  },
];
