/**
 * The tools on a fake api and a fake Graph, for the write, status and docs
 * tests: a connection that already holds every manifest scope, a wait that
 * does not sleep, and a confirm card that records when it was shown. Given
 * `choose`, prifly also has the two-way mail card ("choose-mail"), whose
 * every answer is `choose` (null for Cancel). Given `edit`, it has the
 * editable card ("choose-mail-edit") instead: `edit` is the answer, or
 * receives the mail on the card and answers with the choice and the
 * subject and text as the reader left them. Given `confirmEdit`, its
 * confirm card has editable content ("confirm-edit"): the edited text, or a
 * function of the content, or null for Cancel.
 */

import type { ConfirmSpec } from "@prifly/extension-api";
import type { CardMail } from "../common";
import { outlookTools, type ToolDeps } from "../tools";
import { type FakeAccount, fakeApi, fakeConnections } from "./fake-api";
import { fakeGraph } from "./fake-graph";

export const ACCOUNT: FakeAccount = {
  id: "ms-1",
  provider: "microsoft",
  account: "ada@example.com",
  scopes: ["User.Read", "Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"],
};

export const SCOPES = { microsoft: ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"] };

export const isWrite = (call: string) => !call.startsWith("GET ");

/** A two-way mail card that was shown, and how many writes Graph had seen by then. */
export type Chosen = { spec: ConfirmSpec & { mail: CardMail }; writesBefore: number };

/** What the editable card answers: the choice and the mail as the reader left it. */
export type EditAnswer = { choice: string; subject: string; text: string } | null;

export function setup(
  options: {
    accounts?: FakeAccount[];
    deps?: ToolDeps;
    choose?: string | null;
    edit?: EditAnswer | ((mail: CardMail) => EditAnswer);
    confirmEdit?: string | null | ((content: string) => string | null);
  } = {},
) {
  const graph = fakeGraph();
  const accounts = options.accounts ?? [{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }];
  const connections = fakeConnections({ accounts, provider: graph.provider, scopes: SCOPES });
  const fake = fakeApi({ connections: connections.connections });
  /** For each card shown: how many writes Graph had seen by then. */
  const writesBeforeConfirm: number[] = [];
  fake.api.confirm = async (spec) => {
    fake.confirms.push(spec);
    writesBeforeConfirm.push(graph.calls.filter(isWrite).length);
    return fake.answer.confirm;
  };
  const chooses: Chosen[] = [];
  const { choose } = options;
  if (choose !== undefined) {
    Object.assign(fake.api, {
      features: [...(fake.api.features ?? []), "choose-mail"],
      chooseMail: async (spec: Chosen["spec"]) => {
        chooses.push({ spec, writesBefore: graph.calls.filter(isWrite).length });
        return choose;
      },
    });
  }
  const { edit, confirmEdit } = options;
  if (edit !== undefined) {
    Object.assign(fake.api, {
      features: [...(fake.api.features ?? []), "choose-mail-edit"],
      chooseMailEdit: async (spec: Chosen["spec"]) => {
        chooses.push({ spec, writesBefore: graph.calls.filter(isWrite).length });
        return typeof edit === "function" ? edit(spec.mail) : edit;
      },
    });
  }
  if (confirmEdit !== undefined) {
    Object.assign(fake.api, {
      features: [...(fake.api.features ?? []), "confirm-edit"],
      confirmEdit: async (spec: ConfirmSpec & { content: string }) => {
        fake.confirms.push(spec);
        writesBeforeConfirm.push(graph.calls.filter(isWrite).length);
        return typeof confirmEdit === "function" ? confirmEdit(spec.content) : confirmEdit;
      },
    });
  }
  fake.api.tools?.register(
    outlookTools(() => fake.api, { wait: async () => undefined, ...options.deps }),
  );
  return { graph, connections, fake, writesBeforeConfirm, chooses };
}
