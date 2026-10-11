/**
 * A fake Microsoft Graph for the tests: a `FakeProvider` over a small mailbox
 * (two conversations, a png, a pdf and a forwarded .eml). It keeps to the
 * rules the real Graph enforces on the calls graph.ts makes — `$orderby`
 * leading `$filter`, no `$filter`/`$orderby` with `$search`, no `$orderby`
 * with a `conversationId` filter — so a call that breaks them fails here too.
 * It can answer 429 with `Retry-After`, 410 and a 401 claims challenge.
 */

import type { FakeProvider } from "./fake-api";

type Entry = Record<string, unknown>;

const BOB = { name: "Bob Builder", address: "bob@example.com" };
const ADA = { name: "Ada Lovelace", address: "ada@example.com" };
const DANA = { name: "Dana Scully", address: "dana@example.com" };

const who = (person: { name: string; address: string }) => ({ emailAddress: person });

function mail(fields: Entry): Entry {
  const id = String(fields["id"]);
  return {
    isRead: true,
    isDraft: false,
    hasAttachments: false,
    categories: [],
    flag: { flagStatus: "notFlagged" },
    ccRecipients: [],
    webLink: `https://outlook.office.com/mail/id/${id}`,
    ...fields,
  };
}

const body = (content: string) => ({ contentType: "text", content });

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
      toRecipients: [who(ADA)],
      receivedDateTime: "2026-10-10T07:00:00Z",
      hasAttachments: true,
      flag: {
        flagStatus: "flagged",
        dueDateTime: { dateTime: "2026-10-12T00:00:00.0000000", timeZone: "UTC" },
      },
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

/** A scripted failure for the next requests. */
export type Fault =
  | { kind: "throttle"; status: 429 | 503; retryAfter: string | null; times: number }
  | { kind: "claims" };

export type FakeGraph = {
  provider: FakeProvider;
  messages: Entry[];
  /** Every request's method and path with query, in order. */
  calls: string[];
  /** Set to make the next request(s) fail. */
  fault: Fault | null;
  /** The delta link the fake treats as expired (410 syncStateNotFound). */
  expiredDelta: boolean;
};

const error = (status: number, code: string, message: string, headers?: Record<string, string>) =>
  Response.json({ error: { code, message } }, { status, ...(headers ? { headers } : {}) });

function checkOrder(url: URL): Response | null {
  const filter = url.searchParams.get("$filter");
  const order = url.searchParams.get("$orderby");
  if (url.searchParams.has("$search") && (filter !== null || order !== null)) {
    return error(400, "BadRequest", "$search cannot be combined with $filter or $orderby");
  }
  if (filter !== null && order !== null) {
    const lead = order.split(" ")[0] ?? "";
    if (!filter.startsWith(lead)) {
      return error(400, "InefficientFilter", "The $orderby property must lead the $filter");
    }
    if (filter.includes("conversationId")) {
      return error(400, "InefficientFilter", "conversationId cannot be combined with $orderby");
    }
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
  let out = messages;
  if (filter.includes("isRead eq false")) out = out.filter((m) => m["isRead"] === false);
  if (filter.includes("flagStatus eq 'flagged'")) {
    out = out.filter((m) => (m["flag"] as Entry)["flagStatus"] === "flagged");
  }
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

/** A fake Graph over `mailbox()`; change `fault`, `messages` and `expiredDelta` between calls. */
export function fakeGraph(): FakeGraph {
  let draftCount = 0;
  const graph: FakeGraph = {
    messages: mailbox(),
    calls: [],
    fault: null,
    expiredDelta: false,
    provider: async (request) => {
      const url = new URL(request.url);
      const path = url.pathname.replace(/^\/v1\.0/, "");
      graph.calls.push(`${request.method} ${path}${url.search}`);
      return await route(request, url, path);
    },
  };

  const route = async (request: Request, url: URL, path: string): Promise<Response> => {
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
    return await answer(request, url, path);
  };

  const answer = async (request: Request, url: URL, path: string): Promise<Response> => {
    if (path === "/me") {
      return Response.json({ displayName: "Ada Lovelace", mail: "ada@example.com" });
    }
    if (path === "/me/mailFolders/inbox/messages/delta") return delta(url);
    if (path === "/me/mailFolders/inbox/messages" || path === "/me/messages") {
      return Response.json({ value: select(graph.messages, url) });
    }
    const one = /^\/me\/messages\/([^/]+)(?:\/(.*))?$/.exec(path);
    if (one === null) return error(404, "NotFound", `No route ${path}`);
    const id = decodeURIComponent(one[1] ?? "");
    const rest = one[2] ?? "";
    const found = graph.messages.find((m) => m["id"] === id);
    if (found === undefined) return error(404, "ErrorItemNotFound", "The item was not found");
    return await messageAnswer(request, found, rest);
  };

  const messageAnswer = async (request: Request, found: Entry, rest: string): Promise<Response> => {
    if (rest === "" && request.method === "GET") return Response.json(found);
    if (rest === "" && request.method === "PATCH") {
      const patch = (await request.json()) as { body?: Entry };
      Object.assign(found, { body: patch.body });
      return Response.json(found);
    }
    if (rest === "createReply" || rest === "createReplyAll") {
      draftCount++;
      const draft = mail({
        id: `draft${draftCount}`,
        conversationId: found["conversationId"],
        subject: `RE: ${String(found["subject"])}`,
        isDraft: true,
        from: who(ADA),
        toRecipients: [found["from"]],
        receivedDateTime: "2026-10-11T10:00:00Z",
        body: body(String(((await request.json()) as Entry)["comment"])),
      });
      graph.messages.push(draft);
      return Response.json(draft, { status: 201 });
    }
    return attachmentAnswer(String(found["id"]), rest);
  };

  const attachmentAnswer = (id: string, rest: string): Response => {
    const stored = attachmentsOf(id);
    if (rest === "attachments") return Response.json({ value: stored.map((s) => s.meta) });
    const one = /^attachments\/([^/]+)(\/\$value)?$/.exec(rest);
    const found = stored.find((s) => s.meta["id"] === one?.[1]);
    if (one === null || found === undefined) return error(404, "NotFound", `No route ${rest}`);
    return one[2] === undefined ? Response.json(found.meta) : new Response(found.bytes);
  };

  const delta = (url: URL): Response => {
    if (graph.expiredDelta && url.searchParams.has("$deltatoken")) {
      return error(410, "syncStateNotFound", "The sync state was not found");
    }
    const base = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta";
    if (url.searchParams.has("$deltatoken")) {
      return Response.json({
        value: [],
        "@odata.deltaLink": `${base}?$deltatoken=T${graph.messages.length}`,
      });
    }
    if (!url.searchParams.has("$skiptoken")) {
      return Response.json({
        value: graph.messages.slice(0, 2),
        "@odata.nextLink": `${base}?$skiptoken=2`,
      });
    }
    return Response.json({
      value: graph.messages.slice(2),
      "@odata.deltaLink": `${base}?$deltatoken=T${graph.messages.length}`,
    });
  };

  return graph;
}
