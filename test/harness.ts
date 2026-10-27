/**
 * The tools on a fake api and a fake Graph, for the write, status and docs
 * tests: a connection that already holds every manifest scope, a wait that
 * does not sleep, and a confirm card that records when it was shown. Given
 * `choose`, prifly also has the two-way mail card ("choose-mail"), whose
 * every answer is `choose` (null for Cancel).
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

export function setup(
  options: { accounts?: FakeAccount[]; deps?: ToolDeps; choose?: string | null } = {},
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
  fake.api.tools?.register(
    outlookTools(() => fake.api, { wait: async () => undefined, ...options.deps }),
  );
  return { graph, connections, fake, writesBeforeConfirm, chooses };
}
