import { expect, test } from "bun:test";
import { makeExtension } from "../extension";
import { SCOPES as MANIFEST_SCOPES } from "../sign-in";
import { type FakeAccount, fakeApi, fakeConnections } from "./fake-api";
import { fakeGraph } from "./fake-graph";
import { ACCOUNT, SCOPES } from "./harness";

/** The extension on a fake api, with every `signIn` it asks for kept. */
function boot(accounts: FakeAccount[]) {
  const graph = fakeGraph();
  const connections = fakeConnections({ accounts, provider: graph.provider, scopes: SCOPES });
  const asked: { scopes: string[]; connection: string | undefined }[] = [];
  const signIn = connections.connections.signIn;
  connections.connections.signIn = (provider, scopes, options) => {
    asked.push({ scopes, connection: options?.connection });
    return signIn(provider, scopes, options);
  };
  const fake = fakeApi({ connections: connections.connections });
  const extension = makeExtension({ wait: async () => undefined, every: () => () => undefined });
  extension.activate(fake.api);
  return { fake, connections, asked, extension };
}

test("the tool asks for exactly the manifest's Microsoft scopes", () => {
  expect(MANIFEST_SCOPES).toEqual(["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"]);
});

test("with no account connected, outlook_sign_in connects a new one and the tools then work", async () => {
  const { fake, asked, extension } = boot([]);
  expect(await fake.callTool("outlook_search", {})).toContain("run outlook_sign_in");
  expect(await fake.callTool("outlook_sign_in", {})).toBe(
    "Signed in to Outlook as microsoft-1@example.com.",
  );
  expect(asked).toEqual([{ scopes: MANIFEST_SCOPES, connection: undefined }]);
  expect(await fake.callTool("outlook_search", {})).not.toContain("outlook_sign_in");
  await extension.refreshStatus();
  expect(fake.statuses.at(-1)).toContain("Connected as microsoft-1@example.com");
});

test("an ended grant is signed in again on the same connection", async () => {
  const { fake, connections, asked } = boot([{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }]);
  connections.endpoint.revoke("ms-1");
  expect(await fake.callTool("outlook_search", {})).toContain("run outlook_sign_in");
  expect(await fake.callTool("outlook_sign_in", {})).toBe(
    "Signed in to Outlook as ada@example.com.",
  );
  expect(asked).toEqual([{ scopes: MANIFEST_SCOPES, connection: "ms-1" }]);
  expect(connections.endpoint.isEnded("ms-1")).toBe(false);
});

test('account "new" adds another account; an unknown one is named, not signed in', async () => {
  const { fake, asked } = boot([{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }]);
  expect(await fake.callTool("outlook_sign_in", { account: "nobody@example.com" })).toBe(
    'No connected Microsoft account matches "nobody@example.com". Connected: ada@example.com. Pass account "new" to add another.',
  );
  expect(asked).toEqual([]);
  expect(await fake.callTool("outlook_sign_in", { account: "new" })).toBe(
    "Signed in to Outlook as microsoft-2@example.com.",
  );
  expect(asked).toEqual([{ scopes: MANIFEST_SCOPES, connection: undefined }]);
});

test("a refused sign-in is plain text, the tenant admin named when Microsoft asks for one", async () => {
  const { fake, connections } = boot([]);
  connections.connections.signIn = async () => ({ ok: false, message: "Sign-in timed out." });
  expect(await fake.callTool("outlook_sign_in", {})).toBe("Not signed in: Sign-in timed out.");
  connections.connections.signIn = async () => ({
    ok: false,
    message: "AADSTS65001: consent required",
  });
  expect(await fake.callTool("outlook_sign_in", {})).toContain("Ask your tenant admin");
  connections.connections.signIn = () => Promise.reject(new Error("no port"));
  expect(await fake.callTool("outlook_sign_in", {})).toBe("Not signed in: no port");
});
