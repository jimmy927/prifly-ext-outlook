/**
 * The tools on a fake api and a fake Graph, for the write, status and docs
 * tests: a connection that already holds every manifest scope, a wait that
 * does not sleep, and a confirm card that records when it was shown.
 */

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

export function setup(options: { accounts?: FakeAccount[]; deps?: ToolDeps } = {}) {
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
  fake.api.tools?.register(
    outlookTools(() => fake.api, { wait: async () => undefined, ...options.deps }),
  );
  return { graph, connections, fake, writesBeforeConfirm };
}
