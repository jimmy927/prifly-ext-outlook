/**
 * `outlook_sign_in`: the tool a session runs when the reader asks to connect
 * Outlook, or a tool said no Microsoft account is connected or a grant ended.
 * It asks prifly to open the reader's browser for every scope the manifest
 * names: on the account in use when there is one (a revived grant, or the
 * scopes it lacks), else on a new account.
 */

import type { ExtensionApi } from "@prifly/extension-api";
import { adminConsentText, isAdminConsentRefusal, pickConnection } from "./auth";
import { accountOf, str, type ToolSpec } from "./common";
import manifest from "./prifly-extension.json";

/** The scopes the manifest asks for; `signIn` accepts nothing else. */
export const SCOPES: string[] = manifest.connections.flatMap((each) =>
  each.provider === "microsoft" ? each.scopes : [],
);

/** Signs in a new Microsoft account, or again for the one in use. */
export async function signIn(api: ExtensionApi, args: Record<string, unknown>): Promise<string> {
  const connections = api.connections;
  if (connections === undefined) {
    return "This prifly cannot use connected accounts: update prifly.";
  }
  const asked = str(args["account"]).trim();
  const existing = asked === "new" ? null : pickConnection(connections, accountOf(api, args));
  if (existing === null && asked !== "" && asked !== "new") {
    const names = connections.list("microsoft").map((each) => each.account);
    if (names.length > 0) {
      return `No connected Microsoft account matches "${asked}". Connected: ${names.join(", ")}. Pass account "new" to add another.`;
    }
  }
  try {
    const answer = await connections.signIn(
      "microsoft",
      SCOPES,
      existing === null ? undefined : { connection: existing.id },
    );
    if (answer.ok) return `Signed in to Outlook as ${answer.account}.`;
    if (isAdminConsentRefusal(answer.message) || answer.adminConsentUrl !== undefined) {
      return `Not signed in. ${adminConsentText(SCOPES.join(" and "), answer.message, answer.adminConsentUrl)}`;
    }
    return `Not signed in: ${answer.message}`;
  } catch (error) {
    return `Not signed in: ${error instanceof Error ? error.message : String(error)}`;
  }
}

export const signInTool: ToolSpec = {
  name: "outlook_sign_in",
  description:
    "Opens the reader's browser to sign in to Microsoft for Outlook: for the account in use when its grant ended or lacks scopes, else a new account. " +
    'account "new" always adds another account. Only when the reader asks, or an Outlook tool said no account is connected or to sign in again.',
  inputSchema: {
    type: "object",
    properties: {
      account: {
        type: "string",
        description:
          'The account\'s address or connection id, or "new" for another account. Left out: the one in use.',
      },
    },
    additionalProperties: false,
  },
  run: async (api, args, _ctx, deps) => {
    const answer = await signIn(api, args);
    deps.onSignedIn?.();
    return answer;
  },
};
