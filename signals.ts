/**
 * What the extension tells the reader without being asked: the launcher lit
 * with the unread count, and a chip on a session whose conversation got a new
 * message. Nothing is ever sent into a session.
 */

import type { Decoration } from "@prifly/extension-api";
import { convoKey, nameOf } from "./conversations";
import { timeLabel } from "./format";
import { boardAccess, cacheFor, type Hub, now } from "./hub";
import type { Mail } from "./model";
import type { Reply } from "./state";

/** The launcher's id in the manifest. */
export const LAUNCHER_ID = "mail";

/** Where the unread count stops being exact. */
export const UNREAD_SHOWN = 100;

/** Whether `mail` was sent by the reader. */
function fromSelf(mail: Mail, self: string[]): boolean {
  return mail.from.address !== "" && self.includes(mail.from.address.toLowerCase());
}

/**
 * Notes the messages that came to conversations a session was started from.
 * An unread one from somebody else is a reply to tell the reader about; one
 * read in Outlook since clears it.
 */
export function noteReplies(hub: Hub, connection: string, changed: Mail[], self: string[]): void {
  for (const mail of changed) {
    const started = hub.state.started[convoKey(mail)];
    if (started === undefined || started.connection !== connection || mail.isDraft) continue;
    if (Date.parse(mail.received) <= Date.parse(started.seen)) continue;
    if (fromSelf(mail, self)) {
      started.seen = mail.received;
    } else if (!mail.isRead) {
      started.reply = {
        from: nameOf(mail.from),
        at: mail.received,
        subject: mail.subject,
        url: mail.webLink,
      };
    } else if (started.reply !== null && mail.received >= started.reply.at) {
      started.reply = null;
      started.seen = mail.received;
    }
  }
}

/** The chip a session shows for a reply: clicking it opens the conversation inside prifly. */
function chip(key: string, reply: Reply, at: Date): Decoration {
  return {
    key,
    icon: "dot",
    label: `New reply from ${reply.from}`,
    tone: "info",
    details: [reply.subject, `Received ${timeLabel(reply.at, at)}`],
    ...(reply.url === "" ? {} : { url: reply.url }),
    launcher: LAUNCHER_ID,
  };
}

/** Puts the chips on the sessions with a reply waiting, and takes the others' away. */
export function renderChips(hub: Hub): void {
  const api = hub.api();
  if (api === null) return;
  const bySession: Record<string, Decoration[]> = {};
  for (const [key, started] of Object.entries(hub.state.started)) {
    if (started.reply !== null) bySession[started.session] = [chip(key, started.reply, now(hub))];
  }
  const signature = JSON.stringify(bySession);
  if (signature === hub.shown.last) return;
  hub.shown.last = signature;
  api.show(bySession, []);
}

/** Lights the launcher with the unread count, or warns that the sign-in ended. */
export function updateFlag(hub: Hub): void {
  const api = hub.api();
  if (api === null) return;
  const access = boardAccess(hub, api);
  if (!access.ok) {
    const account = access.problem.connection?.account;
    const ended = access.problem.kind === "needs-sign-in" && account !== undefined;
    api.flag(LAUNCHER_ID, ended ? { tone: "warning", hint: `Sign in again as ${account}` } : null);
    return;
  }
  const count = cacheFor(hub, access.connection.id).lists.unread?.length ?? 0;
  const hint = count >= UNREAD_SHOWN ? `${UNREAD_SHOWN}+ unread` : `${count} unread`;
  api.flag(LAUNCHER_ID, count > 0 ? { tone: "info", hint } : null);
}
