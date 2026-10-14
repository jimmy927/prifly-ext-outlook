/**
 * The tools that change messages already in the mailbox: move to a folder,
 * delete (to Deleted Items), categories, flags and read state. Each takes one
 * id or a list, reads the messages first, shows the reader every message and
 * the change on one confirm card, and changes nothing when they cancel.
 */

import type { ExtensionApi } from "@prifly/extension-api";
import { failureText } from "./auth";
import {
  ACCOUNT_PROPERTY,
  ask,
  list,
  mailLine,
  type Opened,
  open,
  str,
  type ToolDeps,
  type ToolSpec,
} from "./common";
import { folders, type GraphResult, message, moveMessage, patchMessage } from "./graph";
import type { Json, Mail } from "./model";

const MAX_IDS = 50;

/** One kind of change, as the card and the answer word it. */
type Change = {
  title: string;
  /** "Move 2 messages to Archive". */
  summary: (count: number) => string;
  /** The change for one message, as the card shows it under the message. */
  detail: (mail: Mail) => string;
  apply: (opened: Opened, mail: Mail) => Promise<GraphResult<unknown>>;
  /** What the answer says was done to each. */
  done: string;
};

function plural(count: number): string {
  return count === 1 ? "1 message" : `${count} messages`;
}

/** The ids a call names, from `id` and `ids`; or why they will not do. */
function idsOf(args: Record<string, unknown>): string[] | string {
  const ids = [...new Set([...list(args["ids"]), ...list(args["id"])])];
  if (ids.length === 0) return "Give the id of a message (outlook_search lists them), or ids.";
  if (ids.length > MAX_IDS) return `Not done: at most ${MAX_IDS} messages at a time.`;
  return ids;
}

async function readAll(opened: Opened, ids: string[]): Promise<Mail[] | string> {
  const mails: Mail[] = [];
  for (const id of ids) {
    const found = await message(opened.ctx, id);
    if (!found.ok) return `Not done: ${id}: ${failureText(found.failure, opened.account)}`;
    mails.push(found.value);
  }
  return mails;
}

/** Reads the messages, asks once for all of them, then changes each. */
async function changeMessages(
  api: ExtensionApi,
  opened: Opened,
  ids: string[],
  change: Change,
): Promise<string> {
  const mails = await readAll(opened, ids);
  if (typeof mails === "string") return mails;
  const details = mails.map((mail) => `- ${mailLine(mail)}\n  ${change.detail(mail)}`);
  const refused = await ask(api, {
    title: change.title,
    summary: change.summary(mails.length),
    details: details.join("\n"),
  });
  if (refused !== null) return refused;
  const lines: string[] = [];
  for (const mail of mails) {
    const result = await change.apply(opened, mail);
    lines.push(
      result.ok
        ? `- ${mailLine(mail)}: ${change.done}`
        : `- ${mailLine(mail)}: not changed, ${failureText(result.failure, opened.account)}`,
    );
  }
  return lines.join("\n");
}

/** Outlook's well-known folder names, which Graph takes in place of an id. */
const WELL_KNOWN: Record<string, string> = {
  archive: "Archive",
  deleteditems: "Deleted Items",
  drafts: "Drafts",
  inbox: "Inbox",
  junkemail: "Junk Email",
  outbox: "Outbox",
  sentitems: "Sent Items",
  conversationhistory: "Conversation History",
  scheduled: "Scheduled",
};

const FOLDER_ID = /^[A-Za-z0-9+/=_-]{40,}$/;

/** The folder a destination names: an id, a well-known name, or a folder's path or name. */
async function folderFor(
  opened: Opened,
  destination: string,
): Promise<{ id: string; name: string } | string> {
  const wanted = destination.trim();
  const known = WELL_KNOWN[wanted.toLowerCase().replaceAll(" ", "")];
  if (known !== undefined) return { id: wanted.toLowerCase().replaceAll(" ", ""), name: known };
  const all = await folders(opened.ctx);
  if (!all.ok) return failureText(all.failure, opened.account);
  const lower = wanted.toLowerCase();
  const byId = all.value.find((f) => f.id === wanted);
  const byPath = all.value.find((f) => f.path.toLowerCase() === lower);
  const byName = all.value.filter((f) => f.name.toLowerCase() === lower);
  const found = byId ?? byPath ?? (byName.length === 1 ? byName[0] : undefined);
  if (found !== undefined) return { id: found.id, name: found.path };
  if (byName.length > 1) {
    return `Not done: more than one folder is called ${wanted}: ${byName.map((f) => f.path).join(", ")}. Give its path.`;
  }
  if (FOLDER_ID.test(wanted)) return { id: wanted, name: `the folder with id ${wanted}` };
  return `Not done: there is no folder ${wanted}. outlook_folders lists them.`;
}

async function moveText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const ids = idsOf(args);
  if (typeof ids === "string") return ids;
  const destination = str(args["folder"]);
  if (destination.trim() === "") return "Give the folder to move to.";
  const opened = await open(api, args, ["Mail.ReadWrite"], deps);
  if (typeof opened === "string") return opened;
  const folder = await folderFor(opened, destination);
  if (typeof folder === "string") return folder;
  return changeMessages(api, opened, ids, {
    title: "Move mail",
    summary: (n) => `Move ${plural(n)} to ${folder.name}`,
    detail: () => `→ ${folder.name}`,
    apply: (o, mail) => moveMessage(o.ctx, mail.id, folder.id),
    done: `moved to ${folder.name} (its id changes when it moves)`,
  });
}

async function deleteText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const ids = idsOf(args);
  if (typeof ids === "string") return ids;
  const opened = await open(api, args, ["Mail.ReadWrite"], deps);
  if (typeof opened === "string") return opened;
  return changeMessages(api, opened, ids, {
    title: "Delete mail",
    summary: (n) => `Move ${plural(n)} to Deleted Items`,
    detail: () => "→ Deleted Items (it can be restored from there)",
    apply: (o, mail) => moveMessage(o.ctx, mail.id, "deleteditems"),
    done: "moved to Deleted Items; it is not gone for good and can be restored from there",
  });
}

type CategoryMode = "set" | "add" | "remove";

function nextCategories(current: string[], given: string[], mode: CategoryMode): string[] {
  if (mode === "set") return given;
  if (mode === "add") return [...new Set([...current, ...given])];
  const gone = new Set(given.map((each) => each.toLowerCase()));
  return current.filter((each) => !gone.has(each.toLowerCase()));
}

const shown = (categories: string[]) =>
  categories.length === 0 ? "(none)" : categories.join(", ");

async function categoriesText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const ids = idsOf(args);
  if (typeof ids === "string") return ids;
  const given = list(args["categories"]);
  const modeArg = str(args["mode"]) || "set";
  if (modeArg !== "set" && modeArg !== "add" && modeArg !== "remove") {
    return "mode is set, add or remove.";
  }
  const mode: CategoryMode = modeArg;
  if (given.length === 0 && mode !== "set") return "Give the categories to add or remove.";
  const opened = await open(api, args, ["Mail.ReadWrite"], deps);
  if (typeof opened === "string") return opened;
  const verb = {
    set: "Set the categories of",
    add: "Add categories to",
    remove: "Remove categories from",
  };
  return changeMessages(api, opened, ids, {
    title: "Change categories",
    summary: (n) => `${verb[mode]} ${plural(n)}: ${shown(given)}`,
    detail: (mail) =>
      `categories ${shown(mail.categories)} → ${shown(nextCategories(mail.categories, given, mode))}`,
    apply: (o, mail) =>
      patchMessage(o.ctx, mail.id, { categories: nextCategories(mail.categories, given, mode) }),
    done: "categories changed",
  });
}

type FlagAction = "flag" | "complete" | "clear";

const FLAG_STATUS: Record<FlagAction, string> = {
  flag: "flagged",
  complete: "complete",
  clear: "notFlagged",
};

function flagPatch(action: FlagAction, due: string): Json {
  if (action !== "flag" || due === "") return { flag: { flagStatus: FLAG_STATUS[action] } };
  const at = { dateTime: `${due}T00:00:00`, timeZone: "UTC" };
  return { flag: { flagStatus: "flagged", startDateTime: at, dueDateTime: at } };
}

async function flagText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const ids = idsOf(args);
  if (typeof ids === "string") return ids;
  const actionArg = str(args["action"]) || "flag";
  if (actionArg !== "flag" && actionArg !== "complete" && actionArg !== "clear") {
    return "action is flag, complete or clear.";
  }
  const action: FlagAction = actionArg;
  const due = str(args["due"]).trim();
  if (due !== "" && !/^\d{4}-\d{2}-\d{2}$/.test(due)) return "due is a date, YYYY-MM-DD.";
  const opened = await open(api, args, ["Mail.ReadWrite"], deps);
  if (typeof opened === "string") return opened;
  const label = { flag: "Flag", complete: "Mark as complete", clear: "Clear the flag of" };
  const dueText = action === "flag" && due !== "" ? `, due ${due}` : "";
  return changeMessages(api, opened, ids, {
    title: "Change flags",
    summary: (n) => `${label[action]} ${plural(n)}${dueText}`,
    detail: (mail) => `flag ${mail.flag.status} → ${FLAG_STATUS[action]}${dueText}`,
    apply: (o, mail) => patchMessage(o.ctx, mail.id, flagPatch(action, due)),
    done: `flag ${FLAG_STATUS[action]}${dueText}`,
  });
}

async function markReadText(api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) {
  const ids = idsOf(args);
  if (typeof ids === "string") return ids;
  const read = args["read"] !== false;
  const opened = await open(api, args, ["Mail.ReadWrite"], deps);
  if (typeof opened === "string") return opened;
  const state = read ? "read" : "unread";
  return changeMessages(api, opened, ids, {
    title: `Mark mail as ${state}`,
    summary: (n) => `Mark ${plural(n)} as ${state}`,
    detail: (mail) => `${mail.isRead ? "read" : "unread"} → ${state}`,
    apply: (o, mail) => patchMessage(o.ctx, mail.id, { isRead: read }),
    done: `marked ${state}`,
  });
}

const IDS = {
  id: { type: "string", description: "The id of one message, from outlook_search." },
  ids: { type: "array", items: { type: "string" }, description: "Several ids, at most 50." },
};

const CONFIRMS = "The reader confirms on a card that lists each message and the change first.";

function spec(
  name: string,
  description: string,
  properties: Record<string, unknown>,
  run: (api: ExtensionApi, args: Record<string, unknown>, deps: ToolDeps) => Promise<string>,
): ToolSpec {
  return {
    name,
    description: `${description} ${CONFIRMS}`,
    inputSchema: {
      type: "object",
      properties: { ...IDS, ...properties, account: ACCOUNT_PROPERTY },
      additionalProperties: false,
    },
    run: (api, args, _ctx, deps) => run(api, args, deps),
  };
}

export const changeTools: ToolSpec[] = [
  spec(
    "outlook_move",
    "Move Outlook mail to a folder: a folder id, a path or name from outlook_folders, or a well-known name (inbox, archive, junkemail, deleteditems…). A moved message gets a new id.",
    {
      folder: {
        type: "string",
        description: "Where to: an id, a path, a name or a well-known name.",
      },
    },
    moveText,
  ),
  spec(
    "outlook_delete",
    "Delete Outlook mail: it is moved to Deleted Items, not removed for good, and can be restored from there.",
    {},
    deleteText,
  ),
  spec(
    "outlook_categories",
    "Set, add or remove Outlook categories on mail (names as Outlook shows them; an unknown name is made as a new category).",
    {
      categories: { type: "array", items: { type: "string" } },
      mode: {
        type: "string",
        enum: ["set", "add", "remove"],
        description: "Default set: replaces them.",
      },
    },
    categoriesText,
  ),
  spec(
    "outlook_flag",
    "Flag Outlook mail for follow-up (optionally with a due date), mark the flag complete, or clear it.",
    {
      action: { type: "string", enum: ["flag", "complete", "clear"], description: "Default flag." },
      due: { type: "string", description: "With flag: the due date, YYYY-MM-DD." },
    },
    flagText,
  ),
  spec(
    "outlook_mark_read",
    "Mark Outlook mail as read, or with read false as unread.",
    { read: { type: "boolean", description: "Default true; false marks it unread." } },
    markReadText,
  ),
];
