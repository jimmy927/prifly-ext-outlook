/**
 * What the Graph client hands back: mail and attachments as plain records,
 * read tolerantly from Graph's JSON. Nothing here talks to the network.
 */

/** A JSON object whose fields are not known yet. */
export type Json = Record<string, unknown>;

/** The object `value` is, or an empty one: Graph's answers are read, never trusted. */
export function record(value: unknown): Json {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? (value as Json)
    : {};
}

export function text(value: unknown): string {
  return typeof value === "string" ? value : "";
}

export type Person = { name: string; address: string };

export type Mail = {
  id: string;
  conversationId: string;
  subject: string;
  from: Person;
  to: Person[];
  cc: Person[];
  /** ISO time the mail arrived. */
  received: string;
  isRead: boolean;
  isDraft: boolean;
  hasAttachments: boolean;
  preview: string;
  /** Plain text; empty when the call did not select the body. */
  body: string;
  webLink: string;
  /** "flagged", "complete" or "notFlagged", with the flag's due time when it has one. */
  flag: { status: string; due: string | null };
  categories: string[];
};

export type AttachmentKind = "file" | "item" | "reference";

export type Attachment = {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
  kind: AttachmentKind;
  /** A reference attachment's link (OneDrive and the like); null for the others. */
  link: string | null;
};

function toPerson(value: unknown): Person {
  const address = record(record(value)["emailAddress"]);
  return { name: text(address["name"]), address: text(address["address"]) };
}

function toPeople(value: unknown): Person[] {
  return Array.isArray(value) ? value.map(toPerson) : [];
}

/** "Name <address>", or whichever of the two there is. */
export function personLabel(person: Person): string {
  if (person.name === "" || person.name === person.address) return person.address;
  return person.address === "" ? person.name : `${person.name} <${person.address}>`;
}

export function toMail(raw: unknown): Mail {
  const r = record(raw);
  const flag = record(r["flag"]);
  const due = text(record(flag["dueDateTime"])["dateTime"]);
  return {
    id: text(r["id"]),
    conversationId: text(r["conversationId"]),
    subject: text(r["subject"]),
    from: toPerson(r["from"]),
    to: toPeople(r["toRecipients"]),
    cc: toPeople(r["ccRecipients"]),
    received: text(r["receivedDateTime"]),
    isRead: r["isRead"] === true,
    isDraft: r["isDraft"] === true,
    hasAttachments: r["hasAttachments"] === true,
    preview: text(r["bodyPreview"]),
    body: text(record(r["body"])["content"]),
    webLink: text(r["webLink"]),
    flag: { status: text(flag["flagStatus"]) || "notFlagged", due: due === "" ? null : due },
    categories: Array.isArray(r["categories"]) ? r["categories"].map(text) : [],
  };
}

function attachmentKind(odataType: string): AttachmentKind {
  if (odataType.endsWith("itemAttachment")) return "item";
  if (odataType.endsWith("referenceAttachment")) return "reference";
  return "file";
}

export function toAttachment(raw: unknown): Attachment {
  const r = record(raw);
  const link = text(r["sourceUrl"]);
  return {
    id: text(r["id"]),
    name: text(r["name"]),
    contentType: text(r["contentType"]),
    size: typeof r["size"] === "number" ? r["size"] : 0,
    isInline: r["isInline"] === true,
    kind: attachmentKind(text(r["@odata.type"])),
    link: link === "" ? null : link,
  };
}

/** Newest first. */
export function newestFirst(mails: Mail[]): Mail[] {
  return [...mails].sort((a, b) => Date.parse(b.received) - Date.parse(a.received));
}

/** Oldest first: the order a conversation reads in. */
export function oldestFirst(mails: Mail[]): Mail[] {
  return [...mails].sort((a, b) => Date.parse(a.received) - Date.parse(b.received));
}
