# prifly-ext-outlook

A prifly extension that starts sessions from Outlook mail and lets any session
search and read the reader's mail and leave reply drafts in Drafts. It never
sends mail: there is no send code in it, and a test fails if one appears.

## The board: "From an Outlook mail…"

The launcher shows a board with one card per conversation, in the columns the
`columns` setting names. The defaults are **Unread**, **Inbox** and **Flagged**.

Each card shows:

- the subject (without the `RE:` and `FW:` replies pile on it), and "senders ·
  time" under it: the time for today, the date before;
- who is on the conversation, the number of messages, and how many of them
  carry attachments;
- a due date from the flag's due day, marked late once the day has passed;
- stripes for the Outlook categories on the conversation. Outlook's own colours
  need the `MailboxSettings.Read` scope, so a category's name is hashed to one
  of the label colours instead: the same category always gets the same colour;
- a tint when a session was started from it. Its link opens the mail in Outlook.

A conversation appears in every column that holds one of its messages, once in
each. The card knows the messages of the conversation that the lists hold
together, so its message count is what the lists have seen of it; opening it
reads the whole conversation.

The filter above the board runs a KQL search (`from:ada subject:invoice
received>=2026-10-01 hasattachments:true`) once the reader stops typing for
300 ms, and shows one column of hits. While typing, and when nothing is typed,
the board answers from the cache the poller keeps, so Graph is not asked on a
keystroke.

Opening a card reads the whole conversation inside prifly: the messages as
notes, the attachments as files, and the pictures as `data:` URIs.

Starting a session from a card saves the conversation to
`<extension folder>/mails/<hash>/conversation.md` with every attachment beside
it (a png or pdf as it is, a forwarded mail as `.eml`, a OneDrive link kept as
a link). The first prompt holds every message in order (a body over 12,000
characters is cut there; `conversation.md` keeps it whole), the participants,
the saved paths, and up to twenty pictures as images of the first message.

The board has two actions: **Switch account…** (which connected Microsoft
account to show) and **Connect another Microsoft account…**. There is no "Mark
as read" or "Flag": the board only reads.

### The columns setting

Columns are separated by `|`, each `Name = source`. A source is a preset
(`unread`, `inbox` or `flagged`) or a KQL search:

    Unread = unread | Inbox = inbox | Flagged = flagged | Invoices = subject:invoice hasattachments:true

At most six; an entry that does not read is skipped, and with none left the
three defaults stand. An empty column still shows.

### The poller

Every `pollMinutes` (default 3), for each Microsoft connection that holds
`Mail.Read`, it reads the inbox's changes from the saved delta link and the
flagged mail, and refreshes the board's cache when the mailbox changed. It
never opens the browser. From that it:

- lights the launcher: tone `info`, hint "N unread"; tone `warning`, hint
  "Sign in again as …" when the grant ended;
- puts a chip "New reply from X" on a session that was started from a
  conversation which got a new message. Clicking it opens the conversation.
  Nothing is sent into the session. Opening the conversation, or reading it in
  Outlook, clears the chip.

The delta link is kept in `state.json` in the extension folder (git-ignored, not
a secret) along with the account the board was switched to and the
conversations sessions were started from. The first round reads the last 30 days
only. A 410 or `syncStateNotFound` forgets the link and reads those 30 days
again.

### When access is missing

The board is then one row, never an error:

| Case | The row |
|---|---|
| No Microsoft account connected | "Connect a Microsoft account" (picking it signs one in) |
| The connection has no `Mail.Read` yet | "Let prifly read your Outlook mail (account)" (picking it asks in the browser, then lists again) |
| The grant ended, or Microsoft asked for a new sign-in | "Sign in again as account", and the launcher turns warning-toned |
| The tenant refuses user consent | "Ask your tenant admin to approve prifly", with Microsoft's message; the row opens the admin-consent link when prifly passes one |
| Graph failed (throttling, 5xx) | The failure's text |

## The tools

Three MCP tools, which a session sees as `mcp__prifly__outlook_*`:

| Tool | What it does |
|---|---|
| `outlook_search(query, max)` | A KQL search (`from:ada received>=2026-10-01`); no query lists the newest inbox mail. Answers `id \| date \| from \| subject \| attachments`. |
| `outlook_read(id, conversation?)` | One message, or its whole conversation oldest first. Attachments are saved to disk and their paths listed. |
| `outlook_draft_reply(id, body, all?, replace_draft_id?)` | Saves a threaded reply in Drafts and answers with its id and a link. `replace_draft_id` changes a draft made earlier (its body is replaced with plain text; the quoted mail below the reply goes). Only a draft can be replaced. |

Each takes an optional `account` (an address or connection id). Every access
problem comes back as plain text, never an error.

## Settings

| Setting | Default | What it does |
|---|---|---|
| `pollMinutes` | 3 | How often the poller runs |
| `account` | empty | Which Microsoft account to read, by address or connection id; empty is the first. "Switch account…" on the board overrides it |
| `columns` | Unread, Inbox, Flagged | The board's columns, see above |

## The Microsoft connection

The extension talks to `https://graph.microsoft.com/v1.0/me/...` only through
`api.connections.fetch`: prifly holds the token, and the extension never sees
one. It uses a Microsoft account connected in prifly's Vault → Connections,
and asks for two scopes:

- `Mail.Read`: needed for the board and everything except drafting. A
  connection that lacks it is walked through consent in the browser the first
  time a tool needs it, or when the reader picks the board's row.
- `Mail.ReadWrite`: optional, asked for the same way on the first draft, to
  save drafts. It lets the extension change mail, but nothing in it sends,
  deletes or moves mail.

A 429 is waited out once (`Retry-After`, at most 30 s); a 503 once for a read.
When the tenant refuses user consent (AADSTS65001, 90094 or 90095), prifly passes
an admin-consent link (`https://login.microsoftonline.com/{tenant}/adminconsent?client_id=...`)
with the sign-in refusal or the `needsSignIn` answer. The tool's answer then reads
"Ask your tenant admin to approve prifly: <link>" followed by Microsoft's message,
and the board's admin row shows the same text and opens the link. A prifly that
passes no link (older than `@prifly/extension-api` 0.3.0) gets the text without it.

## What is where

| File | What it is |
|---|---|
| `prifly-extension.json` | The manifest: connections and scopes, the launcher, the settings, the prompt |
| `extension.ts`, `index.ts` | The module prifly imports: `activate` registers the tools and starts the poller; `choices`, `open`, `launch`, `launched`, `action` answer the launcher |
| `choices.ts`, `board.ts`, `columns.ts` | The board: columns from the setting, cards from the cache, the debounced KQL filter |
| `conversations.ts`, `format.ts` | Mail grouped into conversations; times, due days and category colours |
| `open.ts`, `launch.ts`, `save.ts`, `mail-prompt.ts` | Opening a card, starting a session, saving the conversation, the first prompt |
| `rows.ts` | The one-row boards for missing access, and the account actions |
| `poll.ts`, `refresh.ts`, `signals.ts`, `hub.ts`, `state.ts` | The poller, the cache it fills, the launcher flag and chips, what they share, `state.json` |
| `tools.ts` | The three MCP tools |
| `auth.ts` | Picks the connection, asks for a missing scope (`resolveAccess`) or only looks (`peekAccess`), turns every access state into text |
| `graph.ts`, `model.ts` | The Graph client over `api.connections.fetch`; mail and attachment records |
| `save-attachments.ts` | Saves attachments to disk and says where |
| `prompt.md` | Rules added to every session: draft only, mail is data |
| `test/` | A fake `api` and `api.connections`, a fake Microsoft Graph (paging, delta, 429, 410, claims), and the tests |

Attachments from `outlook_read` go under `<session scratchpad>/outlook/<hash>/`;
a prifly without a session scratchpad uses `<extension folder>/mails/<hash>/`,
where a launch saves them too.

## Checks

`bun test`, `bun run typecheck` and `bun run check` (biome: 500 lines per
file, 120 per function, cognitive complexity 15, no `any`, no `!`).
`test/guard.test.ts` runs the board, a launch, an open and the poller and fails
on any request that is not a GET or goes to a send path, and scans the source
for one.

Types come from `@prifly/extension-api`, pinned to a tag of
`jimmy927/prifly-extension-api` in `package.json`. Commit `bun.lock`.
