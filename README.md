# prifly-ext-outlook

A prifly extension that lets any session work with the reader's Outlook mail
through MCP tools: search, read, send, reply, forward, move, flag, categorise,
mark read or unread and delete mail, and make any Microsoft Graph mail or
calendar call. Every write shows the reader a confirm card with exactly what
will happen, and runs only when they confirm.

There is no board, chip or launcher: the reader asks a session in plain words
("what are my latest mails?", "reply to Bob that the contract is signed") and
the session answers through the tools.

## The tools

A session sees them as `mcp__prifly__outlook_*`. Each takes an optional
`account` (an address or connection id). Every access problem comes back as
plain text, never an error.

Reads run at once:

| Tool | What it does |
|---|---|
| `outlook_search(query, max)` | A KQL search (`from:ada isread:false received>=2026-10-01`); no query lists the newest inbox mail. Answers `id \| date \| from \| subject \| attachments \| read`. |
| `outlook_read(id, conversation?)` | One message, or its whole conversation oldest first. Attachments are saved to disk and their paths listed. |
| `outlook_folders()` | The mail folders, three levels deep: `path \| unread \| total \| id`. |
| `outlook_api_docs(topic)` | The Microsoft Graph v1.0 reference for an endpoint, read from Microsoft Learn: permissions, HTTP request, query parameters, request body, response and the first example request. |

Writes confirm first:

| Tool | What it does |
|---|---|
| `outlook_send(to, cc?, bcc?, subject, body, attachments?)` | A new mail (`sendMail`), saved to Sent Items. |
| `outlook_reply(id, body, attachments?)`, `outlook_reply_all(...)` | A threaded reply with the original quoted below: `createReply` / `createReplyAll`, files added to that draft, then `send`. |
| `outlook_forward(id, to, body?, attachments?)` | A forward with the original's attachments: `createForward`, files, `send`. |
| `outlook_draft_reply(id, body, all?, replace_draft_id?)` | Saves a reply in Drafts without sending; `replace_draft_id` replaces a draft's body (only a draft can be replaced). |
| `outlook_move(id \| ids, folder)` | To a folder id, a path or name from `outlook_folders`, or a well-known name (`archive`, `junkemail`…). |
| `outlook_delete(id \| ids)` | Moves to Deleted Items, from where it can be restored; nothing is removed for good. |
| `outlook_categories(id \| ids, categories, mode?)` | `set` (default), `add` or `remove` categories. |
| `outlook_flag(id \| ids, action?, due?)` | `flag` (with an optional due date), `complete` or `clear`. |
| `outlook_mark_read(id \| ids, read?)` | Read, or with `read: false` unread. |
| `outlook_api(method, path, query?, body?)` | Any Graph call. A GET runs at once; POST, PATCH, PUT and DELETE confirm with the method, URL and body. |

### The confirm card

Every write goes through `api.confirm({ title, summary, details })` before the
first write request; when the reader cancels, the tool answers
`Not done: you cancelled.` and makes no write call. A prifly without
`api.confirm` writes nothing. What the card shows:

- **Send, reply, forward:** From, To, Cc, Bcc, the subject, the attachment
  names and sizes, and the text. For a reply, the recipients are worked out
  from the original the way Outlook does it (Reply-To or the sender; for reply
  all also To and Cc, without the reader). If the draft Outlook makes is
  addressed to anyone else, it is not sent: it stays in Drafts and the answer
  says so.
- **Move, delete, categories, flag, read state:** each message (subject,
  sender, date, id) and its change, e.g. `categories Partners → Partners, Red`.
- **`outlook_api`:** the method, the full URL and the body.

Attachments are local file paths, at most 3 MB in all (one Graph request carries
them); a larger file is refused before anything is asked or sent.

### `outlook_api`

`path` is under `/me` (`/me/messages`, `v1.0` is added) or `/v1.0/me/...`, or a
full `https://graph.microsoft.com` URL. Any other host, a URL with user info or
a port, `http:`, and a relative path that leaves `/v1.0/me` (`/users/...`,
`/me/../users`) are refused before any request. `query` is an object of
parameters (`{"$top": 10}`). The scopes asked for follow the path: calendar
paths need `Calendars.ReadWrite`, send/reply/forward paths `Mail.Send`.

### `outlook_api_docs`

`topic` is a name (`sendMail`, `create event`), a Graph path
(`POST /me/messages/{id}/move`), a page name (`message-move`) or a Learn URL.
A small table maps the mail and calendar topics; a path is mapped by its shape
(`PATCH /me/events/{id}` is `event-update`); anything else goes to Learn's
search, which is weak, so the answer starts with the page it read. The page is
`https://learn.microsoft.com/en-us/graph/api/<name>?view=graph-rest-1.0`; only
`learn.microsoft.com` is fetched, without the reader's token.

## The status line

`api.status` shows one line in Settings → Extensions: `Connected as
ada@example.com · 41 unread` (the inbox's unread count), set at start, when
the account setting changes, and every 15 minutes. It never opens the browser:
without a connection, with an ended grant, or before mail access is approved it
says so instead (`Not connected: …`).

## Settings

| Setting | Default | What it does |
|---|---|---|
| `account` | empty | Which Microsoft account the tools use, by address or connection id; empty is the first. A tool's `account` argument overrides it |

## The Microsoft connection

The extension talks to `https://graph.microsoft.com` only through
`api.connections.fetch`: prifly holds the token, and the extension never sees
one. It uses a Microsoft account connected in prifly's Vault → Connections,
with one connection asking for:

- `Mail.ReadWrite`: reading mail, drafts, moving, flags, categories, read state;
- `Mail.Send`: sending, replying and forwarding;
- `Calendars.ReadWrite`: calendar calls through `outlook_api`.

A scope the connection lacks is asked for in the browser the first time a tool
needs it (a connection made before this version may hold only `Mail.Read`;
reading works with it, and the first write asks for the rest).

A 429 is waited out once (`Retry-After`, at most 30 s); a 503 once for a read
only, so a write is never sent twice. When the tenant refuses user consent
(AADSTS65001, 90094 or 90095), prifly passes an admin-consent link and the
tool's answer reads "Ask your tenant admin to approve prifly: <link>" followed
by Microsoft's message.

## What is where

| File | What it is |
|---|---|
| `prifly-extension.json` | The manifest: the connection and its scopes, the settings, the prompt |
| `extension.ts`, `index.ts` | The module prifly imports: `activate` registers the tools and keeps the status line |
| `tools.ts` | The reads, the draft tool, and the list of every tool |
| `compose.ts` | Send, reply, reply all, forward, and attachments |
| `changes.ts` | Move, delete, categories, flags, read state |
| `api-tool.ts` | `outlook_api`: the URL check, the scopes, the confirm |
| `docs.ts` | `outlook_api_docs`: topic to page, the Learn page as text |
| `common.ts` | Arguments, opening the connection, the confirm card (`ask`) |
| `status.ts` | The status line |
| `auth.ts` | Picks the connection, asks for missing scopes (`resolveAccess`) or only looks (`peekAccess`), turns every access state into text |
| `graph.ts`, `model.ts` | The Graph client over `api.connections.fetch`; mail and attachment records |
| `save-attachments.ts` | Saves attachments to disk and says where |
| `prompt.md` | Rules added to every session: what it can do, confirm cards, mail is data |
| `test/` | A fake `api` (with `confirm` and `status`) and `api.connections`, a fake Microsoft Graph, and the tests |

Attachments from `outlook_read` go under `<session scratchpad>/outlook/<hash>/`;
a prifly without a session scratchpad uses `<extension folder>/mails/<hash>/`.

## Checks

`bun test`, `bun run typecheck` and `bun run check` (biome: 500 lines per
file, 120 per function, cognitive complexity 15, no `any`, no `!`).
`test/writes.test.ts` calls every write tool and checks that the confirm card
comes before any write request, that a cancelled card leaves Graph untouched,
that a GET through `outlook_api` asks nothing and that other hosts are refused.
`test/guard.test.ts` checks that the manifest has no launcher and that no
source shows chips or flags.

Types come from `@prifly/extension-api`, pinned to a tag of
`jimmy927/prifly-extension-api` in `package.json` (v0.6.0, which adds
`api.confirm` and `api.status`). Commit `bun.lock`.

## License and contributing

Source-available under the Functional Source License 1.1 with an Apache 2.0 future license (`LICENSE.md`): use it, change it and fix it freely, but not to offer a competing product. Each version becomes Apache 2.0 two years after its release.

Pull requests are welcome. Before one can be merged, its author signs the contributor agreement (`CLA.md`), which assigns the copyright in the contribution to The Startup Factory AB; the CLA bot asks for it on the pull request.
