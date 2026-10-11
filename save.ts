/**
 * Saves a conversation for a session to open: `<dir>/conversation.md` and
 * every attachment beside it, as Trello's `cards/<id>/` does. The pictures
 * are read back from disk to ride along as the first message's images.
 */

import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { Launch } from "@prifly/extension-api";
import type { GraphContext } from "./graph";
import { type AttachmentLines, conversationMarkdown } from "./mail-prompt";
import type { Mail } from "./model";
import { type SavedFile, saveAttachmentFiles } from "./save-attachments";

/** Launch carries at most this many pictures. */
export const MAX_IMAGES = 20;

/** A picture bigger than this stays a file on disk. */
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;

export type SavedConversation = {
  dir: string;
  /** Null when the file could not be written. */
  markdownPath: string | null;
  files: AttachmentLines;
  saved: SavedFile[];
};

/** Saves every attachment, then the conversation text naming where each went. */
export async function saveConversation(
  graph: GraphContext,
  account: string,
  dir: string,
  mails: Mail[],
): Promise<SavedConversation> {
  const files: AttachmentLines = new Map();
  const saved: SavedFile[] = [];
  const counter = { next: 1 };
  for (const mail of mails) {
    const result = await saveAttachmentFiles(graph, account, dir, mail, counter);
    files.set(mail.id, result.lines);
    saved.push(...result.saved);
  }
  const path = join(dir, "conversation.md");
  let markdownPath: string | null = path;
  try {
    await mkdir(dir, { recursive: true });
    await writeFile(path, conversationMarkdown(mails, files));
  } catch {
    markdownPath = null;
  }
  return { dir, markdownPath, files, saved };
}

type ImageType = NonNullable<Launch["images"]>[number]["mediaType"];

/** The picture type a file is, by its content type or else its extension; null for anything else. */
export function imageType(contentType: string, name: string): ImageType | null {
  const type = contentType.toLowerCase();
  if (
    type === "image/png" ||
    type === "image/jpeg" ||
    type === "image/gif" ||
    type === "image/webp"
  ) {
    return type;
  }
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  return ext === "webp" ? "image/webp" : null;
}

/** The saved pictures as `Launch.images`, up to twenty, none over 5 MB. */
export async function pictureImages(saved: SavedFile[]): Promise<NonNullable<Launch["images"]>> {
  const images: NonNullable<Launch["images"]> = [];
  for (const file of saved) {
    const mediaType = imageType(file.contentType, file.name);
    if (mediaType === null || images.length >= MAX_IMAGES) continue;
    try {
      if ((await stat(file.path)).size > MAX_IMAGE_BYTES) continue;
      images.push({ mediaType, data: (await readFile(file.path)).toString("base64") });
    } catch {
      // A picture that cannot be read stays a path in the prompt.
    }
  }
  return images;
}
