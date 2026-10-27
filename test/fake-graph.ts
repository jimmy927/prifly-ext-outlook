/**
 * A fake Microsoft Graph for the tests: a `FakeProvider` over a small mailbox
 * (two conversations, a png, a pdf and a forwarded .eml) and a few folders.
 * It keeps to the rules the real Graph enforces on the calls graph.ts makes —
 * `$orderby` leading `$filter`, no `$filter`/`$orderby` with `$search`, no
 * `$orderby` with a `conversationId` filter — so a call that breaks them
 * fails here too. It answers the writes the tools make (sendMail, a new draft,
 * createReply, createReplyAll, createForward, attachments, send, move, PATCH), keeping what
 * was sent in `sent`, and can answer 429 with `Retry-After` and a 401 claims
 * challenge.
 */

import type { FakeProvider } from "./fake-api";

type Entry = Record<string, unknown>;

export const BOB = { name: "Bob Builder", address: "bob@example.com" };
export const ADA = { name: "Ada Lovelace", address: "ada@example.com" };
export const DANA = { name: "Dana Scully", address: "dana@example.com" };

export const who = (person: { name: string; address: string }) => ({ emailAddress: person });

export function mail(fields: Entry): Entry {
  const id = String(fields["id"]);
  return {
    isRead: true,
    isDraft: false,
    hasAttachments: false,
    categories: [],
    flag: { flagStatus: "notFlagged" },
    ccRecipients: [],
    replyTo: [],
    webLink: `https://outlook.office.com/mail/id/${id}`,
    ...fields,
  };
}

export const body = (content: string) => ({ contentType: "text", content });

/** Two conversations: A (Bob and Ada, a png and a pdf) and B (Dana, flagged, a forwarded mail and a link). */
export function mailbox(): Entry[] {
  return [
    mail({
      id: "m1",
      conversationId: "convA",
      subject: "Contract for review",
      from: who(BOB),
      toRecipients: [who(ADA)],
      receivedDateTime: "2026-10-09T08:00:00Z",
      hasAttachments: true,
      bodyPreview: "Please see the contract",
      body: body("Hi Ada, please see the contract and the logo."),
    }),
    mail({
      id: "m2",
      conversationId: "convA",
      subject: "RE: Contract for review",
      from: who(ADA),
      toRecipients: [who(BOB)],
      ccRecipients: [who(DANA)],
      receivedDateTime: "2026-10-09T09:30:00Z",
      bodyPreview: "Looks fine",
      body: body("Looks fine to me, Bob."),
    }),
    mail({
      id: "m3",
      conversationId: "convB",
      subject: "Partner Center notice",
      from: who(DANA),
      toRecipients: [who(ADA), who(BOB)],
      receivedDateTime: "2026-10-10T07:00:00Z",
      hasAttachments: true,
      flag: {
        flagStatus: "flagged",
        dueDateTime: { dateTime: "2026-10-12T00:00:00.0000000", timeZone: "UTC" },
      },
      categories: ["Partners"],
      bodyPreview: "Forwarding the notice",
      body: body("Forwarding the notice from Microsoft."),
    }),
    mail({
      id: "m4",
      conversationId: "convB",
      subject: "RE: Partner Center notice",
      from: who(DANA),
      toRecipients: [who(ADA)],
      receivedDateTime: "2026-10-11T06:00:00Z",
      isRead: false,
      hasAttachments: true,
      bodyPreview: "Any news?",
      body: body("Any news on the notice?"),
    }),
  ];
}

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const PDF = new TextEncoder().encode("%PDF-1.4 fake contract");
const EML = new TextEncoder().encode(
  "From: dana@example.com\r\nSubject: Original notice\r\n\r\nHello",
);

type StoredAttachment = { meta: Entry; bytes: Uint8Array };

function attachmentsOf(messageId: string): StoredAttachment[] {
  const file = (id: string, name: string, contentType: string, bytes: Uint8Array) => ({
    meta: {
      "@odata.type": "#microsoft.graph.fileAttachment",
      id,
      name,
      contentType,
      size: bytes.length,
      isInline: false,
    },
    bytes,
  });
  if (messageId === "m1") {
    return [
      file("a1", "logo.png", "image/png", PNG),
      file("a2", "contract.pdf", "application/pdf", PDF),
    ];
  }
  if (messageId === "m3") {
    return [
      {
        meta: {
          "@odata.type": "#microsoft.graph.itemAttachment",
          id: "a3",
          name: "Original notice",
          contentType: "message/rfc822",
          size: EML.length,
          isInline: false,
        },
        bytes: EML,
      },
    ];
  }
  if (messageId === "m4") {
    return [
      {
        meta: {
          "@odata.type": "#microsoft.graph.referenceAttachment",
          id: "a4",
          name: "Notice.docx",
          contentType: "text/html",
          size: 0,
          isInline: false,
          sourceUrl: "https://onedrive.example.com/notice.docx",
        },
        bytes: new Uint8Array(),
      },
    ];
  }
  return [];
}

/** The folders: Inbox (with Clients under it), Archive and a second "Clients" at the top. */
export const FOLDERS: Entry[] = [
  {
    id: "f-inbox",
    displayName: "Inbox",
    childFolderCount: 1,
    unreadItemCount: 1,
    totalItemCount: 4,
  },
  {
    id: "f-archive",
    displayName: "Archive",
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 9,
  },
  {
    id: "f-projects",
    displayName: "Projects",
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 2,
  },
];
const CHILDREN: Entry[] = [
  {
    id: "f-clients",
    displayName: "Clients",
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 3,
  },
];

/** A scripted failure for the next requests. */
export type Fault =
  | { kind: "throttle"; status: 429 | 503; retryAfter: string | null; times: number }
  | { kind: "claims" };

export type FakeGraph = {
  provider: FakeProvider;
  messages: Entry[];
  /** Every request's method and path with query, in order. */
  calls: string[];
  /** The JSON body of every request that had one, by its call. */
  bodies: { call: string; body: Entry }[];
  /** Mail that went out: a `sendMail` message, or a draft that was sent. */
  sent: Entry[];
  /** Files added to drafts, by draft id. */
  attached: Map<string, Entry[]>;
  /** Set to make the next request(s) fail. */
  fault: Fault | null;
  /** Recipients the next created draft gets instead of the ones worked out from the original. */
  draftTo: Entry[] | null;
};

const error = (status: number, code: string, message: string, headers?: Record<string, string>) =>
  Response.json({ error: { code, message } }, { status, ...(headers ? { headers } : {}) });

function checkOrder(url: URL): Response | null {
  const filter = url.searchParams.get("$filter");
  const order = url.searchParams.get("$orderby");
  if (url.searchParams.has("$search") && (filter !== null || order !== null)) {
    return error(400, "BadRequest", "$search cannot be combined with $filter or $orderby");
  }
  if (filter !== null && order !== null && filter.includes("conversationId")) {
    return error(400, "InefficientFilter", "conversationId cannot be combined with $orderby");
  }
  return null;
}

function select(messages: Entry[], url: URL): Entry[] {
  const filter = url.searchParams.get("$filter") ?? "";
  const term = (url.searchParams.get("$search") ?? "")
    .replace(/^"|"$/g, "")
    .replace(/\\"/g, '"')
    .toLowerCase();
  const conversation = /conversationId eq '([^']*)'/.exec(filter)?.[1];
  let out = messages.filter((m) => m["isDraft"] !== true);
  if (conversation !== undefined) out = out.filter((m) => m["conversationId"] === conversation);
  if (term !== "") {
    const needle = term.replace(/^[a-z]+:/, "");
    out = out.filter((m) =>
      JSON.stringify([m["subject"], m["body"], m["from"]]).toLowerCase().includes(needle),
    );
  }
  if (url.searchParams.get("$orderby") === "receivedDateTime desc") {
    out = [...out].sort((a, b) =>
      String(b["receivedDateTime"]).localeCompare(String(a["receivedDateTime"])),
    );
  }
  const top = Number(url.searchParams.get("$top") ?? "1000");
  return out.slice(0, top);
}

type Person = { emailAddress: { address: string } };
const addressOf = (entry: unknown) => (entry as Person).emailAddress.address;
const notAda = (entry: unknown) => addressOf(entry) !== ADA.address;

/** A new mail as Graph saves it in Drafts from `POST /me/messages`, the `n`th draft made. */
const savedDraft = (json: Entry, n: number): Entry =>
  mail({
    ...json,
    id: `draft${n}`,
    conversationId: `conv-draft${n}`,
    isDraft: true,
    from: who(ADA),
    receivedDateTime: "2026-10-11T10:00:00Z",
  });

/** A fake Graph over `mailbox()`; change `fault` and `messages` between calls. */
export function fakeGraph(): FakeGraph {
  let draftCount = 0;
  const graph: FakeGraph = {
    messages: mailbox(),
    calls: [],
    bodies: [],
    sent: [],
    attached: new Map(),
    fault: null,
    draftTo: null,
    provider: async (request) => {
      const url = new URL(request.url);
      const path = url.pathname.replace(/^\/v1\.0/, "");
      const call = `${request.method} ${path}${url.search}`;
      graph.calls.push(call);
      const text = request.method === "GET" ? "" : await request.text();
      const json = text === "" ? {} : (JSON.parse(text) as Entry);
      if (text !== "") graph.bodies.push({ call, body: json });
      return route(request.method, url, path, json);
    },
  };

  const route = (method: string, url: URL, path: string, json: Entry): Response => {
    const fault = graph.fault;
    if (fault?.kind === "claims") {
      return new Response("", {
        status: 401,
        headers: {
          "WWW-Authenticate": 'Bearer error="insufficient_claims", claims="eyJhY2Nlc3MifQ=="',
        },
      });
    }
    if (fault?.kind === "throttle" && fault.times > 0) {
      fault.times--;
      const headers: Record<string, string> =
        fault.retryAfter === null ? {} : { "Retry-After": fault.retryAfter };
      return error(fault.status, "TooManyRequests", "Slow down", headers);
    }
    const ordered = checkOrder(url);
    if (ordered !== null) return ordered;
    return answer(method, url, path, json);
  };

  /** The routes that are not one message, by "METHOD path". */
  const fixed: Record<string, (url: URL, json: Entry) => Response> = {
    "GET /me": () => Response.json({ displayName: "Ada Lovelace", mail: "ada@example.com" }),
    "GET /me/mailFolders/inbox": () =>
      Response.json({
        id: "f-inbox",
        unreadItemCount: graph.messages.filter((m) => m["isRead"] === false).length,
      }),
    "GET /me/mailFolders": () => Response.json({ value: FOLDERS }),
    "GET /me/mailFolders/f-inbox/childFolders": () => Response.json({ value: CHILDREN }),
    "POST /me/sendMail": (_url, json) => {
      graph.sent.push(json["message"] as Entry);
      return new Response(null, { status: 202 });
    },
    "POST /me/messages": (_url, json) => {
      const draft = savedDraft(json, ++draftCount);
      graph.messages.push(draft);
      return Response.json(draft, { status: 201 });
    },
    "GET /me/mailFolders/inbox/messages": (url) =>
      Response.json({ value: select(graph.messages, url) }),
    "GET /me/messages": (url) => Response.json({ value: select(graph.messages, url) }),
  };

  const answer = (method: string, url: URL, path: string, json: Entry): Response => {
    const route = fixed[`${method} ${path}`];
    if (route !== undefined) return route(url, json);
    const one = /^\/me\/messages\/([^/]+)(?:\/(.*))?$/.exec(path);
    if (one === null) return error(404, "NotFound", `No route ${path}`);
    const id = decodeURIComponent(one[1] ?? "");
    const found = graph.messages.find((m) => m["id"] === id);
    if (found === undefined) return error(404, "ErrorItemNotFound", "The item was not found");
    return messageAnswer(method, found, one[2] ?? "", json);
  };

  const createDraft = (found: Entry, rest: string, json: Entry): Response => {
    draftCount++;
    const original = found;
    const replyTo =
      (original["replyTo"] as Entry[]).length > 0 ? original["replyTo"] : [original["from"]];
    const sender = replyTo as Entry[];
    const all = rest === "createReplyAll";
    const to =
      rest === "createForward"
        ? (json["toRecipients"] as Entry[])
        : all
          ? [...sender, ...(original["toRecipients"] as Entry[]).filter(notAda)]
          : sender;
    const cc = all ? (original["ccRecipients"] as Entry[]).filter(notAda) : [];
    const prefix = rest === "createForward" ? "FW" : "RE";
    const draft = mail({
      id: `draft${draftCount}`,
      conversationId: original["conversationId"],
      subject: `${prefix}: ${String(original["subject"]).replace(/^RE: /, "")}`,
      isDraft: true,
      from: who(ADA),
      toRecipients: graph.draftTo ?? to,
      ccRecipients: graph.draftTo === null ? cc : [],
      receivedDateTime: "2026-10-11T10:00:00Z",
      body: body(String(json["comment"])),
    });
    graph.messages.push(draft);
    return Response.json(draft, { status: 201 });
  };

  /** What can be done to one message, by "METHOD rest-of-path". */
  const onMessage: Record<string, (found: Entry, json: Entry) => Response> = {
    "GET ": (found) => Response.json(found),
    "PATCH ": (found, json) => {
      Object.assign(found, json);
      return Response.json(found);
    },
    "POST createReply": (found, json) => createDraft(found, "createReply", json),
    "POST createReplyAll": (found, json) => createDraft(found, "createReplyAll", json),
    "POST createForward": (found, json) => createDraft(found, "createForward", json),
    "POST attachments": (found, json) => {
      const list = graph.attached.get(String(found["id"])) ?? [];
      list.push(json);
      graph.attached.set(String(found["id"]), list);
      return Response.json({ id: `att${list.length}`, ...json }, { status: 201 });
    },
    "POST send": (found) => {
      found["isDraft"] = false;
      graph.sent.push(found);
      return new Response(null, { status: 202 });
    },
    "POST move": (found, json) => {
      found["parentFolderId"] = json["destinationId"];
      return Response.json({ ...found, id: `${String(found["id"])}-moved` }, { status: 201 });
    },
  };

  const messageAnswer = (method: string, found: Entry, rest: string, json: Entry): Response => {
    const route = onMessage[`${method} ${rest}`];
    return route === undefined ? attachmentAnswer(String(found["id"]), rest) : route(found, json);
  };

  const attachmentAnswer = (id: string, rest: string): Response => {
    const stored = attachmentsOf(id);
    if (rest === "attachments") return Response.json({ value: stored.map((s) => s.meta) });
    const one = /^attachments\/([^/]+)(\/\$value)?$/.exec(rest);
    const found = stored.find((s) => s.meta["id"] === one?.[1]);
    if (one === null || found === undefined) return error(404, "NotFound", `No route ${rest}`);
    return one[2] === undefined ? Response.json(found.meta) : new Response(found.bytes);
  };

  return graph;
}
