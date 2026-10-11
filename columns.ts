/**
 * The board's columns, as the `columns` setting writes them: entries
 * separated by `|`, each `Name = source`, where the source is a preset
 * (`unread`, `inbox`, `flagged`) or a KQL search (`from:dana hasattachments:true`).
 * Anything that does not read is skipped; with nothing left the three
 * defaults stand.
 */

export type Preset = "unread" | "inbox" | "flagged";

export type ColumnSource = { kind: "preset"; preset: Preset } | { kind: "kql"; kql: string };

export type ColumnSpec = { id: string; name: string; source: ColumnSource };

export const DEFAULT_COLUMNS = "Unread = unread | Inbox = inbox | Flagged = flagged";

const MAX_COLUMNS = 6;

function isPreset(value: string): value is Preset {
  return value === "unread" || value === "inbox" || value === "flagged";
}

function sourceOf(text: string): ColumnSource {
  const word = text.toLowerCase();
  return isPreset(word) ? { kind: "preset", preset: word } : { kind: "kql", kql: text };
}

function parseEntry(entry: string): { name: string; source: ColumnSource } | null {
  const text = entry.trim();
  if (text === "") return null;
  const at = text.indexOf("=");
  const name = at < 0 ? text : text.slice(0, at).trim();
  const source = at < 0 ? text : text.slice(at + 1).trim();
  return name === "" || source === "" ? null : { name, source: sourceOf(source) };
}

/** The columns the setting names, in order. */
export function parseColumns(setting: string): ColumnSpec[] {
  const found: ColumnSpec[] = [];
  for (const entry of setting.split("|")) {
    const parsed = parseEntry(entry);
    if (parsed !== null && found.length < MAX_COLUMNS)
      found.push({ id: `c${found.length}`, ...parsed });
  }
  return found.length > 0 ? found : parseColumns(DEFAULT_COLUMNS);
}
