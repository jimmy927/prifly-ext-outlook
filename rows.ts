/**
 * The board when there is no mailbox to show: one row saying what is missing
 * and what picking it does, plus the account actions. Picking a row (or an
 * action) reaches `launch` with an empty prompt, so it signs in or switches
 * and the chooser lists again.
 */

import type { ExtensionApi, Launch, LaunchBoard, LaunchChoice } from "@prifly/extension-api";
import {
  adminConsentText,
  isAdminConsentRefusal,
  NO_CONNECTION_TEXT,
  type Problem,
  pickConnection,
} from "./auth";
import {
  ACTION_PREFIX,
  accessRow,
  CONNECT_ACTION,
  ROW_PREFIX,
  rowsBoard,
  switchAction,
} from "./board";
import { cacheFor, type Hub, persist } from "./hub";
import { updateFlag } from "./signals";

/** The account actions of every board that has a connection to switch from. */
export function boardActions(api: ExtensionApi): LaunchChoice[] {
  const names = api.connections?.list("microsoft").map((each) => each.account) ?? [];
  return names.length === 0 ? [] : [switchAction(names), CONNECT_ACTION];
}

/** The admin-consent row: the text, and the link the tenant's admin opens when there is one. */
function adminRow(refused: { text: string; url?: string }): LaunchChoice {
  const row = accessRow("admin", "Ask your tenant admin to approve prifly", refused.text);
  if (refused.url !== undefined) row.url = refused.url;
  return row;
}

function rowFor(hub: Hub, problem: Problem): LaunchChoice {
  const account = problem.connection?.account ?? "";
  switch (problem.kind) {
    case "no-connection":
      return accessRow("connect", "Connect a Microsoft account", NO_CONNECTION_TEXT);
    case "needs-sign-in": {
      const refused = hub.refused.get(problem.connection?.id ?? "");
      return refused === undefined
        ? accessRow("sign-in", `Sign in again as ${account}`, problem.text)
        : adminRow(refused);
    }
    case "consent-needed": {
      const refused = hub.refused.get(problem.connection?.id ?? "");
      return refused === undefined
        ? accessRow(
            "grant",
            problem.text,
            "Opens your browser to approve Mail.Read. prifly reads mail and never sends it.",
          )
        : adminRow(refused);
    }
    default:
      return accessRow("error", problem.text);
  }
}

/** The one-row board for an access problem. */
export function accessBoard(hub: Hub, api: ExtensionApi, problem: Problem): LaunchBoard {
  return rowsBoard("Outlook", [rowFor(hub, problem)], boardActions(api));
}

/** The one-row board for a Graph failure, with the text it came back with. */
export function failureBoard(api: ExtensionApi, text: string): LaunchBoard {
  return rowsBoard("Outlook", [accessRow("error", text)], boardActions(api));
}

function said(message: string): Launch {
  return { prompt: "", message };
}

/** Asks for Mail.Read on `connection`, or for a new account when it is null. */
async function signInTo(hub: Hub, api: ExtensionApi, connection: string | null): Promise<Launch> {
  if (api.connections === undefined) return said("This prifly cannot use connected accounts.");
  try {
    const result = await api.connections.signIn(
      "microsoft",
      ["Mail.Read"],
      connection === null ? {} : { connection },
    );
    if (!result.ok) {
      const url = result.adminConsentUrl;
      if (connection === null || !(isAdminConsentRefusal(result.message) || url !== undefined)) {
        hub.refused.delete(connection ?? "");
        return said(`Not connected: ${result.message}`);
      }
      const text = adminConsentText("Mail.Read", result.message, url);
      hub.refused.set(connection, url === undefined ? { text } : { text, url });
      return said(text);
    }
    hub.refused.delete(result.connection);
    hub.ended.delete(result.connection);
    hub.state.account = result.connection;
    await persist(hub);
    return said(`Connected ${result.account}.`);
  } catch (error) {
    return said(`Sign-in did not start: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** Switches the board to the account `input` names. */
async function switchTo(hub: Hub, api: ExtensionApi, input: string): Promise<Launch> {
  const found = api.connections === undefined ? null : pickConnection(api.connections, input);
  if (found === null || input.trim() === "") {
    const names = api.connections?.list("microsoft").map((each) => each.account) ?? [];
    return said(
      `No connected Microsoft account matches "${input.trim()}". Connected: ${names.join(", ")}.`,
    );
  }
  hub.state.account = found.id;
  await persist(hub);
  updateFlag(hub);
  return said(`Showing ${found.account}.`);
}

/** What picking a row or an account action does; `connection` is the board's account. */
export async function rowLaunch(
  hub: Hub,
  api: ExtensionApi,
  key: string,
  input: string,
  connection: string | null,
): Promise<Launch> {
  let launch: Launch;
  if (key === `${ACTION_PREFIX}switch`) launch = await switchTo(hub, api, input);
  else if (key === `${ACTION_PREFIX}connect` || key === `${ROW_PREFIX}connect`) {
    launch = await signInTo(hub, api, null);
  } else if (key === `${ROW_PREFIX}error`) {
    if (connection !== null) hub.caches.delete(connection);
    launch = said("Reading your mail again.");
  } else launch = await signInTo(hub, api, connection);
  if (connection !== null) cacheFor(hub, connection).triedAt = 0;
  updateFlag(hub);
  return launch;
}
