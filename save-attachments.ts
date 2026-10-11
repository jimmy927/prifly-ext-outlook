/**
 * Saves a message's attachments to disk for a session to open, and says where
 * each went. A file attachment is saved as its bytes, a forwarded mail (an
 * item attachment) as an `.eml`, and a reference attachment (a OneDrive link)
 * is kept as its link only.
 */

import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { ExtensionApi, ExtensionToolContext } from "@prifly/extension-api";
import { failureText } from "./auth";
import { attachments, downloadAttachment, type GraphContext } from "./graph";
import type { Attachment, Mail } from "./model";

/** Bigger attachments are listed, not downloaded. */
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;

/** A short stable name for a conversation's folder. */
export function folderHash(conversationId: string): string {
  return createHash("sha256").update(conversationId).digest("hex").slice(0, 12);
}

/**
 * Where a conversation's attachments go: `<scratchpad>/outlook/<hash>/` when
 * the host gives the call a scratchpad, else `<extension folder>/mails/<hash>/`.
 * `ctx.scratchpad` is newer than the first api version, hence the check.
 */
export function attachmentDir(
  api: ExtensionApi,
  ctx: ExtensionToolContext,
  conversationId: string,
): string {
  const hash = folderHash(conversationId);
  if ("scratchpad" in ctx && typeof ctx.scratchpad === "string") {
    return join(ctx.scratchpad, "outlook", hash);
  }
  return join(api.folder, "mails", hash);
}

/** A file name that is safe on every disk: no separators, no control characters. */
export function safeName(name: string): string {
  const cleaned = name
    .replace(/[^\p{L}\p{N}._ -]/gu, "_")
    .replace(/^\.+/, "_")
    .trim();
  return cleaned === "" ? "attachment" : cleaned.slice(0, 120);
}

function kb(bytes: number): string {
  return `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

/** Why an attachment is not downloaded, or null when it should be. */
function skipReason(attachment: Attachment): string | null {
  if (attachment.kind === "reference") return `link ${attachment.link ?? "(none given)"}`;
  if (attachment.isInline) return "inline image, not saved";
  if (attachment.size > MAX_ATTACHMENT_BYTES) {
    return "too large to save here (over 25 MB), open the mail in Outlook";
  }
  return null;
}

/** Writes the file; null when it landed, else why not. */
async function writeOut(dir: string, path: string, bytes: Uint8Array): Promise<string | null> {
  try {
    await mkdir(dir, { recursive: true });
    await writeFile(path, bytes);
    return null;
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
}

/**
 * One line per attachment of `mail`: the saved path, or the link, or why it
 * was not saved. `counter` numbers files across a whole conversation so equal
 * names never collide.
 */
export async function saveAttachments(
  graph: GraphContext,
  account: string,
  dir: string,
  mail: Mail,
  counter: { next: number },
): Promise<string[]> {
  return (await saveAttachmentFiles(graph, account, dir, mail, counter)).lines;
}

/** A file that was written to disk. */
export type SavedFile = { path: string; name: string; contentType: string; size: number };

/** `saveAttachments`, also saying which files landed on disk (the board's launch reads the pictures). */
export async function saveAttachmentFiles(
  graph: GraphContext,
  account: string,
  dir: string,
  mail: Mail,
  counter: { next: number },
): Promise<{ lines: string[]; saved: SavedFile[] }> {
  const lines: string[] = [];
  const saved: SavedFile[] = [];
  if (!mail.hasAttachments) return { lines, saved };
  const listed = await attachments(graph, mail.id);
  if (!listed.ok) {
    lines.push(`- attachments could not be listed: ${failureText(listed.failure, account)}`);
    return { lines, saved };
  }
  for (const attachment of listed.value) {
    const label = `${attachment.name} (${attachment.contentType || attachment.kind}, ${kb(attachment.size)})`;
    const skipped = skipReason(attachment);
    if (skipped !== null) {
      lines.push(`- ${label}: ${skipped}`);
      continue;
    }
    const bytes = await downloadAttachment(graph, mail.id, attachment.id);
    if (!bytes.ok) {
      lines.push(`- ${label}: not saved, ${failureText(bytes.failure, account)}`);
      continue;
    }
    const base = safeName(attachment.name);
    const named = attachment.kind === "item" && !base.endsWith(".eml") ? `${base}.eml` : base;
    const path = join(dir, `${counter.next++}-${named}`);
    const failed = await writeOut(dir, path, bytes.value);
    lines.push(`- ${label}: ${failed === null ? `saved to ${path}` : `not saved, ${failed}`}`);
    if (failed === null) {
      saved.push({ path, name: named, contentType: attachment.contentType, size: attachment.size });
    }
  }
  return { lines, saved };
}
