/**
 * Conversations out of lists of mail. A list (unread, the inbox, a search)
 * holds some of a conversation's messages; everything the lists hold together
 * is what the card knows of it: its newest message, how many messages, who
 * is on it, whether a flag is due.
 */

import { localDay } from "./format";
import { type Mail, oldestFirst, type Person } from "./model";

export type Convo = {
  /** The conversation's id; a mail without one stands alone under its own id. */
  key: string;
  /** Every message known, oldest first. */
  mails: Mail[];
  newest: Mail;
};

export function convoKey(mail: Mail): string {
  return mail.conversationId === "" ? mail.id : mail.conversationId;
}

/** Every mail the lists hold, once each, grouped by conversation, oldest first. */
export function knownConversations(lists: Mail[][]): Map<string, Mail[]> {
  const byId = new Map<string, Mail>();
  for (const list of lists) for (const mail of list) byId.set(mail.id, mail);
  const grouped = new Map<string, Mail[]>();
  for (const mail of byId.values()) {
    if (mail.isDraft) continue;
    const key = convoKey(mail);
    grouped.set(key, [...(grouped.get(key) ?? []), mail]);
  }
  for (const [key, mails] of grouped) grouped.set(key, oldestFirst(mails));
  return grouped;
}

/** The conversations of `mails`, newest first, each with everything `known` has of it. */
export function conversationsOf(mails: Mail[], known: Map<string, Mail[]>): Convo[] {
  const seen = new Set<string>();
  const found: Convo[] = [];
  for (const mail of mails) {
    const key = convoKey(mail);
    if (seen.has(key) || mail.isDraft) continue;
    seen.add(key);
    const all = known.get(key) ?? [mail];
    const newest = all[all.length - 1] ?? mail;
    found.push({ key, mails: all, newest });
  }
  return found.sort((a, b) => Date.parse(b.newest.received) - Date.parse(a.newest.received));
}

export function nameOf(person: Person): string {
  return person.name !== "" ? person.name : person.address;
}

/** Whether `person` is one of the reader's own addresses (lower case). */
export function isSelf(person: Person, self: string[]): boolean {
  return person.address !== "" && self.includes(person.address.toLowerCase());
}

/** Who sent the messages, oldest sender first; "A, B +2" past three. */
export function sendersLine(convo: Convo): string {
  const names: string[] = [];
  for (const mail of convo.mails) {
    const name = nameOf(mail.from);
    if (name !== "" && !names.includes(name)) names.push(name);
  }
  const shown = names.slice(0, 3).join(", ");
  return names.length > 3 ? `${shown} +${names.length - 3}` : shown;
}

/** Everyone on the conversation but the reader, the newest senders first. */
export function participants(convo: Convo, self: string[]): string[] {
  const names: string[] = [];
  for (const mail of [...convo.mails].reverse()) {
    for (const person of [mail.from, ...mail.to, ...mail.cc]) {
      const name = nameOf(person);
      if (name !== "" && !isSelf(person, self) && !names.includes(name)) names.push(name);
    }
  }
  return names;
}

/** The categories of every message, once each, in the order they appear. */
export function categoriesOf(convo: Convo): string[] {
  const names: string[] = [];
  for (const mail of convo.mails) {
    for (const name of mail.categories) if (name !== "" && !names.includes(name)) names.push(name);
  }
  return names;
}

/** The earliest due day (`YYYY-MM-DD`) among the conversation's flagged messages, or null. */
export function dueDay(convo: Convo): string | null {
  const days = convo.mails
    .filter((mail) => mail.flag.status === "flagged" && mail.flag.due !== null)
    .map((mail) => (mail.flag.due ?? "").slice(0, 10))
    .sort();
  return days[0] ?? null;
}

/** A due day before today's date is late; today's is not. */
export function isLate(day: string, now: Date): boolean {
  return day < localDay(now);
}

/** How many of the messages carry attachments. */
export function attachmentCount(convo: Convo): number {
  return convo.mails.filter((mail) => mail.hasAttachments).length;
}

/** The subject without the "RE: " and "FW: " the replies pile on it. */
export function topic(subject: string): string {
  const bare = subject.replace(/^\s*((re|fw|fwd|sv|vs|aw|wg)\s*:\s*)+/i, "").trim();
  return bare === "" ? subject.trim() || "(no subject)" : bare;
}
