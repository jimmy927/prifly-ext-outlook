/**
 * A card opened inside prifly: the whole conversation as a `LaunchItem`.
 * Messages are its notes, attachments its files, and pictures come as `data:`
 * URIs so the window can show them.
 */

import type { LaunchItem, LaunchItemFile } from "@prifly/extension-api";
import { failureText } from "./auth";
import {
  type Convo,
  categoriesOf,
  dueDay,
  isLate,
  nameOf,
  participants,
  topic,
} from "./conversations";
import { categoryColour, dayLabel } from "./format";
import { attachments, conversation, downloadAttachment, type GraphContext } from "./graph";
import { boardAccess, type Hub, now, persist, selfOf } from "./hub";
import { type Attachment, type Mail, oldestFirst } from "./model";
import { imageType, MAX_IMAGES } from "./save";
import { LAUNCHER_ID, renderChips } from "./signals";

/** A picture bigger than this is a link in Outlook, not a `data:` URI in the window. */
const MAX_PICTURE_BYTES = 2 * 1024 * 1024;

function plain(markdown: string): LaunchItem {
  return { title: "Outlook", markdown };
}

/** A picture as a `data:` URI, or null when it is not a small picture file or could not be fetched. */
async function pictureData(
  ctx: GraphContext,
  mail: Mail,
  attachment: Attachment,
): Promise<string | null> {
  const type = imageType(attachment.contentType, attachment.name);
  if (type === null || attachment.kind !== "file") return null;
  const bytes = await downloadAttachment(ctx, mail.id, attachment.id);
  if (!bytes.ok || bytes.value.length > MAX_PICTURE_BYTES) return null;
  return `data:${type};base64,${Buffer.from(bytes.value).toString("base64")}`;
}

/** One message's files: pictures with their data, the rest as a link to open elsewhere. */
async function filesOfMail(
  ctx: GraphContext,
  mail: Mail,
  pictures: { left: number },
): Promise<LaunchItemFile[]> {
  const listed = mail.hasAttachments ? await attachments(ctx, mail.id) : null;
  const files: LaunchItemFile[] = [];
  for (const attachment of listed?.ok ? listed.value : []) {
    if (attachment.isInline) continue;
    const file: LaunchItemFile = { id: `${mail.id}:${attachment.id}`, name: attachment.name };
    const link = attachment.link ?? mail.webLink;
    if (link !== "") file.url = link;
    const data = pictures.left > 0 ? await pictureData(ctx, mail, attachment) : null;
    if (data !== null) {
      file.data = data;
      pictures.left--;
    }
    files.push(file);
  }
  return files;
}

async function filesOf(ctx: GraphContext, mails: Mail[]): Promise<LaunchItemFile[]> {
  const pictures = { left: MAX_IMAGES };
  const files: LaunchItemFile[] = [];
  for (const mail of mails) files.push(...(await filesOfMail(ctx, mail, pictures)));
  return files;
}

/** The reader has looked at the conversation: whatever reply was waiting is dealt with. */
async function markSeen(hub: Hub, convo: Convo): Promise<void> {
  const started = hub.state.started[convo.key];
  if (started === undefined) return;
  started.seen = convo.newest.received;
  started.reply = null;
  await persist(hub);
  renderChips(hub);
}

function itemOf(convo: Convo, files: LaunchItemFile[], self: string[], at: Date): LaunchItem {
  const { newest } = convo;
  const due = dueDay(convo);
  const people = participants(convo, self);
  const count = convo.mails.length;
  return {
    title: topic(newest.subject),
    ...(newest.webLink === "" ? {} : { url: newest.webLink }),
    labels: categoriesOf(convo).map((name) => ({ colour: categoryColour(name), name })),
    people,
    ...(due === null ? {} : { due: dayLabel(due), dueLate: isLate(due, at) }),
    markdown: `${count} message${count === 1 ? "" : "s"} between ${people.join(", ") || nameOf(newest.from)}.`,
    files,
    notes: convo.mails.map((mail) => ({
      by: nameOf(mail.from),
      at: mail.received,
      text: mail.body.trim(),
    })),
  };
}

export async function open(hub: Hub, launchId: string, key: string): Promise<LaunchItem> {
  await hub.ready;
  const api = hub.api();
  if (api === null || launchId !== LAUNCHER_ID)
    return plain("The Outlook extension is not running.");
  const access = boardAccess(hub, api);
  if (!access.ok) return plain(access.problem.text);
  const all = await conversation(access.ctx, key);
  if (!all.ok) return plain(failureText(all.failure, access.connection.account));
  const mails = oldestFirst(all.value.filter((mail) => !mail.isDraft));
  const newest = mails[mails.length - 1];
  if (newest === undefined) return plain("That conversation is not in the mailbox any more.");
  const convo: Convo = { key, mails, newest };
  await markSeen(hub, convo);
  const files = await filesOf(access.ctx, mails);
  return itemOf(convo, files, selfOf(hub, access.connection), now(hub));
}
