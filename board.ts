/**
 * The "From an Outlook mail…" board as data: one card per conversation in
 * each column that holds it, and the rows that stand in for the board when
 * access is missing. Nothing here talks to the network.
 */

import type { LaunchBoard, LaunchChoice } from "@prifly/extension-api";
import {
  attachmentCount,
  type Convo,
  categoriesOf,
  conversationsOf,
  dueDay,
  isLate,
  knownConversations,
  participants,
  sendersLine,
  topic,
} from "./conversations";
import { categoryColour, dayLabel, timeLabel } from "./format";
import type { Mail } from "./model";

/** A column and the mail its source found. */
export type BoardColumn = { id: string; name: string; mails: Mail[] };

export type BoardInput = {
  columns: BoardColumn[];
  /** Every mail the lists hold, so a card knows more of its conversation than one column does. */
  lists: Mail[][];
  /** Conversation keys a session was started from. */
  started: Set<string>;
  /** The reader's own addresses, lower case. */
  self: string[];
  now: Date;
  actions: LaunchChoice[];
};

const MAX_STRIPES = 4;
const MAX_PEOPLE = 6;

export function cardOf(convo: Convo, group: string, input: Omit<BoardInput, "columns" | "lists">) {
  const { newest } = convo;
  const time = timeLabel(newest.received, input.now);
  const senders = sendersLine(convo);
  const due = dueDay(convo);
  const attachments = attachmentCount(convo);
  const card: LaunchChoice = {
    key: convo.key,
    title: topic(newest.subject),
    group,
    detail: [senders, time].filter((part) => part !== "").join(" · "),
    people: participants(convo, input.self).slice(0, MAX_PEOPLE),
    badges: {
      comments: convo.mails.length,
      ...(attachments > 0 ? { attachments } : {}),
      ...(due === null ? {} : { due: dayLabel(due), dueLate: isLate(due, input.now) }),
    },
    stripes: categoriesOf(convo).slice(0, MAX_STRIPES).map(categoryColour),
    started: input.started.has(convo.key),
  };
  if (newest.webLink !== "") card.url = newest.webLink;
  return card;
}

/** The cards of one list of mail, newest conversation first. */
export function cardsOf(
  mails: Mail[],
  group: string,
  input: Omit<BoardInput, "columns" | "lists">,
  known: Map<string, Mail[]>,
): LaunchChoice[] {
  return conversationsOf(mails, known).map((convo) => cardOf(convo, group, input));
}

/** The board: every column, even an empty one. */
export function buildBoard(input: BoardInput): LaunchBoard {
  const known = knownConversations(input.lists);
  const items = input.columns.flatMap((column) => cardsOf(column.mails, column.id, input, known));
  return {
    columns: input.columns.map(({ id, name }) => ({ id, name })),
    items,
    actions: input.actions,
  };
}

/** Whether the words typed all appear in the card's text. */
export function matches(card: LaunchChoice, query: string): boolean {
  const hay = [card.title, card.detail, ...(card.people ?? [])].join(" ").toLowerCase();
  return query
    .toLowerCase()
    .split(/\s+/)
    .filter((word) => word !== "")
    .every((word) => hay.includes(word));
}

/** The board narrowed to the cards that match what was typed, from what is in hand. */
export function narrowed(board: LaunchBoard, query: string): LaunchBoard {
  return { ...board, items: board.items.filter((card) => matches(card, query)) };
}

/** What a row of the access board does when picked; see `launcher.ts`. */
export type RowKind = "grant" | "sign-in" | "admin" | "connect" | "error";

export const ROW_PREFIX = "row:";
export const ACTION_PREFIX = "action:";

export function accessRow(kind: RowKind, title: string, detail?: string): LaunchChoice {
  const row: LaunchChoice = {
    key: `${ROW_PREFIX}${kind}`,
    title,
    group: "outlook",
    tone: "warning",
  };
  if (detail !== undefined) row.detail = detail;
  return row;
}

/** A board of one column holding only `rows`, for when there is nothing to list. */
export function rowsBoard(
  name: string,
  rows: LaunchChoice[],
  actions: LaunchChoice[],
): LaunchBoard {
  return { columns: [{ id: "outlook", name }], items: rows, actions };
}

/** "Switch account…" asks which of the connected accounts to show. */
export function switchAction(accounts: string[]): LaunchChoice {
  return {
    key: `${ACTION_PREFIX}switch`,
    title: "Switch account…",
    input: { title: "Which Microsoft account?", placeholder: accounts.join(", ") },
  };
}

export const CONNECT_ACTION: LaunchChoice = {
  key: `${ACTION_PREFIX}connect`,
  title: "Connect another Microsoft account…",
};
