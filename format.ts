/**
 * How the board writes times and colours: a time for today, a date before;
 * a category's name as one of the label colours. Local time, as the reader
 * reads their own clock.
 */

import type { LabelColour } from "@prifly/extension-api";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const pad = (n: number) => String(n).padStart(2, "0");

/** "2026-10-11" for the local day of `date`. */
export function localDay(date: Date): string {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** "12 Oct" from "2026-10-12"; the input back when it is not a day. */
export function dayLabel(day: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(day);
  const month = MONTHS[Number(match?.[2]) - 1];
  return match === null || month === undefined ? day : `${Number(match[3])} ${month}`;
}

/** "09:30" for a mail that came today, "9 Oct" for an earlier one; "" for a time that is not one. */
export function timeLabel(iso: string, now: Date): string {
  const at = new Date(iso);
  if (Number.isNaN(at.getTime())) return "";
  if (localDay(at) === localDay(now)) return `${pad(at.getHours())}:${pad(at.getMinutes())}`;
  return dayLabel(localDay(at));
}

/**
 * The colours a category can wear. Outlook's own colours need
 * MailboxSettings.Read, one more consent; hashing the name gives the same
 * category the same colour every time without it. No grey: it reads as "none".
 */
const PALETTE: LabelColour[] = ["red", "orange", "yellow", "green", "blue", "purple", "pink"];

export function categoryColour(name: string): LabelColour {
  let hash = 2166136261;
  for (const char of name.trim().toLowerCase()) {
    hash = Math.imul(hash ^ (char.codePointAt(0) ?? 0), 16777619) >>> 0;
  }
  return PALETTE[hash % PALETTE.length] ?? "blue";
}
