/**
 * Which Microsoft connection to use, and what to say when access is missing.
 * Every way access can fail becomes a `Problem` with plain text (the tools
 * return it as their answer; the board will show it as a row) — nothing here
 * throws.
 */

import type { ExtensionConnection, ExtensionConnectionsApi } from "@prifly/extension-api";
import type { GraphContext, GraphFailure, Wait } from "./graph";

export type MailScope = "Mail.Read" | "Mail.ReadWrite";

export type ProblemKind =
  /** This prifly has no `api.connections`. */
  | "no-host-support"
  | "no-connection"
  /** The `account` setting or argument matches no connection. */
  | "unknown-account"
  | "needs-sign-in"
  /** The connection lacks the scope and nobody has asked the reader for it yet (`peekAccess`). */
  | "consent-needed"
  /** The reader did not approve the scope in the browser (or it did not take). */
  | "scope-missing"
  /** The tenant refuses user consent: its admin has to approve prifly. */
  | "admin-consent";

export type Problem = {
  kind: ProblemKind;
  /** What to tell the reader. */
  text: string;
  /** The connection the problem is about, when one was picked. */
  connection: ExtensionConnection | null;
};

export type Access =
  | { ok: true; ctx: GraphContext; connection: ExtensionConnection }
  | { ok: false; problem: Problem };

export const NO_CONNECTION_TEXT =
  "No Microsoft account is connected: connect one in prifly's Vault → Connections.";

/** Microsoft's codes for "the user may not consent here": needs an admin (65001, 90094, 90095). */
const ADMIN_CONSENT = /AADSTS(65001|90094|90095)/;

export function isAdminConsentRefusal(message: string): boolean {
  return ADMIN_CONSENT.test(message);
}

/**
 * What to tell the reader when the tenant needs an admin to approve prifly.
 * `adminConsentUrl` is the link the admin opens, when the host passed one.
 */
export function adminConsentText(
  scope: MailScope,
  providerMessage: string,
  adminConsentUrl?: string,
): string {
  if (adminConsentUrl === undefined) {
    return (
      `Ask your tenant admin to approve prifly for ${scope}: your organisation's Microsoft 365 ` +
      `admin has to grant it before prifly can read this mailbox. Microsoft said: ${providerMessage}`
    );
  }
  return (
    `Ask your tenant admin to approve prifly: ${adminConsentUrl} Your organisation's Microsoft ` +
    `365 admin has to grant ${scope} before prifly can read this mailbox. ` +
    `Microsoft said: ${providerMessage}`
  );
}

export function needsSignInText(account: string): string {
  return `The Outlook grant for ${account} ended: sign in again in Vault → Connections.`;
}

/** The scope's bare name: Graph may list it as `https://graph.microsoft.com/Mail.Read`. */
function bare(scope: string): string {
  return (scope.split("/").pop() ?? scope).toLowerCase();
}

/** Whether the connection holds `scope`; `Mail.ReadWrite` covers `Mail.Read`. */
export function hasScope(connection: ExtensionConnection, scope: MailScope): boolean {
  const held = connection.scopes.map(bare);
  const wanted = scope.toLowerCase();
  return held.includes(wanted) || (scope === "Mail.Read" && held.includes("mail.readwrite"));
}

/** The connection `account` names (its id or address), else the first Microsoft one. */
export function pickConnection(
  connections: ExtensionConnectionsApi,
  account: string,
): ExtensionConnection | null {
  const all = connections.list("microsoft");
  const wanted = account.trim().toLowerCase();
  if (wanted === "") return all[0] ?? null;
  return (
    all.find((each) => each.id.toLowerCase() === wanted || each.account.toLowerCase() === wanted) ??
    null
  );
}

/** A failed Graph call as the text a tool answers with. */
export function failureText(failure: GraphFailure, account: string): string {
  switch (failure.kind) {
    case "needs-sign-in":
      return failure.adminConsentUrl === undefined
        ? needsSignInText(account)
        : `${needsSignInText(account)} Ask your tenant admin to approve prifly: ${failure.adminConsentUrl}`;
    case "unauthorized":
      return needsSignInText(account);
    case "forbidden":
      return `Microsoft refused access to ${account}'s mail (403): ${failure.message}`;
    case "gone":
      return "That item is gone or its sync state expired (410).";
    case "throttled": {
      const wait =
        failure.retryAfterSeconds === null ? "a minute" : `${failure.retryAfterSeconds} s`;
      return `Microsoft Graph is throttling requests (${failure.status}); try again in ${wait}.`;
    }
    case "http":
      return `Microsoft Graph answered ${failure.status}: ${failure.message}`;
    case "refused":
      return `prifly refused the Graph request: ${failure.message}`;
  }
}

function problem(kind: ProblemKind, text: string, connection: ExtensionConnection | null): Access {
  return { ok: false, problem: { kind, text, connection } };
}

/** Asks the reader, in their browser, for `scope` on the connection; null when it is held after. */
async function consent(
  connections: ExtensionConnectionsApi,
  connection: ExtensionConnection,
  scope: MailScope,
): Promise<Problem | null> {
  let result: Awaited<ReturnType<ExtensionConnectionsApi["signIn"]>>;
  try {
    result = await connections.signIn("microsoft", [scope], { connection: connection.id });
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    return { kind: "scope-missing", connection, text: `${scope} was not requested: ${reason}` };
  }
  if (!result.ok) {
    const adminConsent =
      isAdminConsentRefusal(result.message) || result.adminConsentUrl !== undefined;
    const text = adminConsent
      ? adminConsentText(scope, result.message, result.adminConsentUrl)
      : `${scope} was not granted for ${connection.account}: ${result.message} ` +
        "Approve prifly in the browser window it opens, then run this again.";
    return {
      kind: adminConsent ? "admin-consent" : "scope-missing",
      connection,
      text,
    };
  }
  const after = pickConnection(connections, connection.id);
  if (after !== null && hasScope(after, scope)) return null;
  return {
    kind: "scope-missing",
    connection: after ?? connection,
    text:
      `${scope} is still missing for ${connection.account} after sign-in: approve prifly in ` +
      "the browser window it opens, then run this again.",
  };
}

function contextFor(
  connections: ExtensionConnectionsApi,
  connection: ExtensionConnection,
  wait: Wait | undefined,
): Access {
  const ctx: GraphContext = {
    connections,
    connection: connection.id,
    ...(wait === undefined ? {} : { wait }),
  };
  return { ok: true, ctx, connection };
}

/** What the board row and the sentence say when `scope` has not been approved yet. */
export function consentNeededText(scope: MailScope, account: string): string {
  return scope === "Mail.Read"
    ? `Let prifly read your Outlook mail (${account})`
    : `Let prifly save reply drafts for ${account}`;
}

/**
 * Like `resolveAccess`, but never opens the browser: a connection that lacks
 * `scope` is a `consent-needed` problem. For what runs without the reader
 * asking: the poller and the board's listing.
 */
export function peekAccess(
  connections: ExtensionConnectionsApi | undefined,
  options: { account: string; scope: MailScope; wait?: Wait },
): Access {
  if (connections === undefined) {
    return problem(
      "no-host-support",
      "This prifly cannot use connected accounts: update prifly.",
      null,
    );
  }
  if (pickConnection(connections, "") === null) {
    return problem("no-connection", NO_CONNECTION_TEXT, null);
  }
  const connection = pickConnection(connections, options.account);
  if (connection === null) {
    const names = connections.list("microsoft").map((each) => each.account);
    return problem(
      "unknown-account",
      `No connected Microsoft account matches "${options.account}". Connected: ${names.join(", ")}.`,
      null,
    );
  }
  if (connection.state === "needs-sign-in") {
    return problem("needs-sign-in", needsSignInText(connection.account), connection);
  }
  if (!hasScope(connection, options.scope)) {
    return problem(
      "consent-needed",
      consentNeededText(options.scope, connection.account),
      connection,
    );
  }
  return contextFor(connections, connection, options.wait);
}

/**
 * A ready Graph context for the connection `account` names (empty: the
 * first), asking the reader for `scope` first when the connection lacks it.
 * That is the incremental consent: `signIn` with the connection's id, which
 * opens the browser and resolves when they finish.
 */
export async function resolveAccess(
  connections: ExtensionConnectionsApi | undefined,
  options: { account: string; scope: MailScope; wait?: Wait },
): Promise<Access> {
  const peeked = peekAccess(connections, options);
  const missing = peeked.ok ? null : peeked.problem.connection;
  if (connections === undefined || missing === null) return peeked;
  if (peeked.ok || peeked.problem.kind !== "consent-needed") return peeked;
  const refused = await consent(connections, missing, options.scope);
  if (refused !== null) return { ok: false, problem: refused };
  return contextFor(connections, pickConnection(connections, missing.id) ?? missing, options.wait);
}
