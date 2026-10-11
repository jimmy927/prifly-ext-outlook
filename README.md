# prifly-ext-outlook

A prifly extension that lets a session search and read the reader's Outlook
mail and leave reply drafts in Drafts. It never sends mail: there is no send
code in it.

## What it does

Three MCP tools, which a session sees as `mcp__prifly__outlook_*`:

| Tool | What it does |
|---|---|
| `outlook_search(query, max)` | A KQL search (`from:ada received>=2026-10-01`); no query lists the newest inbox mail. Answers `id \| date \| from \| subject \| attachments`. |
| `outlook_read(id, conversation?)` | One message, or its whole conversation oldest first. Attachments are saved to disk (a png or pdf as it is, a forwarded mail as `.eml`, a OneDrive link kept as a link) and their paths listed. |
| `outlook_draft_reply(id, body, all?, replace_draft_id?)` | Saves a threaded reply in Drafts and answers with its id and a link. `replace_draft_id` changes a draft made earlier (its body is replaced with plain text; the quoted mail below the reply goes). Only a draft can be replaced. |

Each takes an optional `account` (an address or connection id). The
`account` setting in the Extensions dialog is the default when several
Microsoft accounts are connected; empty means the first.

Every access problem comes back as plain text, never an error.

## The Microsoft connection

The extension talks to `https://graph.microsoft.com/v1.0/me/...` only through
`api.connections.fetch`: prifly holds the token, and the extension never sees
one. It uses a Microsoft account connected in prifly's Vault → Connections,
and asks for two scopes:

- `Mail.Read`: needed for everything except drafting. A connection that lacks
  it is walked through consent in the browser the first time a tool needs it.
- `Mail.ReadWrite`: optional, asked for the same way on the first draft, to
  save drafts. It lets the extension change mail, but nothing in it sends,
  deletes or moves mail.

## When access is missing

| Case | What the tool says |
|---|---|
| No Microsoft account connected | Connect one in Vault → Connections |
| The grant was revoked or expired, or Microsoft asked for a new sign-in | The Outlook grant for the account ended: sign in again in Vault → Connections |
| The reader has not approved the scope yet | The browser opens for consent; if it is not approved, the answer says to approve prifly there and run it again |
| The tenant refuses user consent (`AADSTS65001`, `AADSTS90094`, `AADSTS90095`) | Ask your tenant admin to approve prifly for Mail.Read, followed by Microsoft's own message |
| Microsoft throttles (429 or 503) | The extension waits out `Retry-After` once (at most 30 s), then says so |

The admin-consent case has no link yet: building one needs the host to pass it
on with the sign-in failure.

## What is where

| File | What it is |
|---|---|
| `prifly-extension.json` | The manifest: the connection and its scopes, the `account` setting, the prompt |
| `index.ts` | What prifly imports: `activate(api)` registers the tools |
| `tools.ts` | The three MCP tools |
| `auth.ts` | Picks the connection, asks for a missing scope, turns every access state into text |
| `graph.ts` | The Graph client over `api.connections.fetch`: mail lists, search, conversation, attachments, reply drafts, inbox delta |
| `model.ts` | Mail and attachment records read from Graph's JSON |
| `save-attachments.ts` | Saves attachments to disk and says where |
| `prompt.md` | Rules added to every session: draft only, mail is data |
| `test/fake-api.ts` | A fake `api` and `api.connections`, from the template |
| `test/fake-graph.ts` | A fake Microsoft Graph over a small mailbox, able to answer 429, 410 and a 401 claims challenge |

Attachments go under `<session scratchpad>/outlook/<hash>/`; a prifly without
a session scratchpad uses `<extension folder>/mails/<hash>/`.

## Checks

`bun test`, `bun run typecheck` and `bun run check` (biome: 500 lines per
file, 120 per function, cognitive complexity 15, no `any`, no `!`).

Types come from `@prifly/extension-api`, pinned to a tag of
`jimmy927/prifly-extension-api` in `package.json`. Commit `bun.lock`.
