/**
 * Starting a session from a card: the conversation is saved to
 * `<extension folder>/mails/<hash>/conversation.md` with its attachments
 * beside it, and the first prompt holds every message. Picking an access row
 * or an account action reaches here too and answers with an empty prompt.
 */

import { join } from "node:path";
import type { ExtensionApi, Launch } from "@prifly/extension-api";
import { failureText } from "./auth";
import { ACTION_PREFIX, ROW_PREFIX } from "./board";
import { topic } from "./conversations";
import { conversation } from "./graph";
import { boardAccess, type Hub, now, persist } from "./hub";
import { firstPrompt } from "./mail-prompt";
import { oldestFirst } from "./model";
import { rowLaunch } from "./rows";
import { pictureImages, saveConversation } from "./save";
import { folderHash } from "./save-attachments";
import { LAUNCHER_ID } from "./signals";

const NAME_LIMIT = 60;

function said(message: string): Launch {
  return { prompt: "", message };
}

async function launchConversation(hub: Hub, api: ExtensionApi, key: string): Promise<Launch> {
  const access = boardAccess(hub, api);
  if (!access.ok) return said(access.problem.text);
  const account = access.connection.account;
  const all = await conversation(access.ctx, key);
  if (!all.ok) return said(failureText(all.failure, account));
  const mails = oldestFirst(all.value.filter((mail) => !mail.isDraft));
  const newest = mails[mails.length - 1];
  if (newest === undefined) return said("That conversation is not in the mailbox any more.");
  const dir = join(api.folder, "mails", folderHash(key));
  const saved = await saveConversation(access.ctx, account, dir, mails);
  const images = await pictureImages(saved.saved);
  hub.launching.set(key, newest.received);
  return {
    prompt: firstPrompt({
      mails,
      account,
      markdownPath: saved.markdownPath,
      files: saved.files,
      pictures: images.length,
    }),
    name: `Mail: ${topic(newest.subject)}`.slice(0, NAME_LIMIT),
    ...(images.length > 0 ? { images } : {}),
  };
}

export async function launch(
  hub: Hub,
  launchId: string,
  key: string,
  input: string,
): Promise<Launch> {
  await hub.ready;
  const api = hub.api();
  if (api === null || launchId !== LAUNCHER_ID)
    return said("The Outlook extension is not running.");
  if (key.startsWith(ROW_PREFIX) || key.startsWith(ACTION_PREFIX)) {
    const access = boardAccess(hub, api);
    const connection = access.ok ? access.connection.id : (access.problem.connection?.id ?? null);
    return rowLaunch(hub, api, key, input, connection);
  }
  return launchConversation(hub, api, key);
}

/** The launch became a session: the card is tinted and the poller watches the conversation. */
export async function launched(
  hub: Hub,
  launchId: string,
  key: string,
  sessionId: string,
): Promise<void> {
  await hub.ready;
  const api = hub.api();
  if (api === null || launchId !== LAUNCHER_ID) return;
  const access = boardAccess(hub, api);
  const connection = access.ok ? access.connection.id : (access.problem.connection?.id ?? "");
  const seen = hub.launching.get(key) ?? now(hub).toISOString();
  hub.launching.delete(key);
  hub.state.started[key] = { session: sessionId, connection, seen, reply: null };
  await persist(hub);
}
