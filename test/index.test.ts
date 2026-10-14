import { afterEach, beforeEach, expect, test } from "bun:test";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { ExtensionModule } from "@prifly/extension-api";
import * as extension from "../index";
import { outlookTools } from "../tools";
import { type FakeAccount, fakeApi, fakeConnections } from "./fake-api";
import { fakeGraph } from "./fake-graph";

// The module is what prifly imports: it must fit the contract as a whole.
const module: ExtensionModule = extension;

const ACCOUNT: FakeAccount = {
  id: "ms-1",
  provider: "microsoft",
  account: "ada@example.com",
  scopes: ["User.Read", "Mail.Read"],
};

// The manifest's scopes: Mail.Read is not among them, so a missing Mail.Read is asked as Mail.ReadWrite.
const SCOPES = { microsoft: ["Mail.ReadWrite", "Mail.Send", "Calendars.ReadWrite"] };

let folder = "";
let stop: (() => void) | undefined;
const waits: number[] = [];

beforeEach(async () => {
  folder = await mkdtemp(join(tmpdir(), "outlook-test-"));
  waits.length = 0;
});

afterEach(async () => {
  stop?.();
  await rm(folder, { recursive: true, force: true });
});

/** The extension on a fake api, its tools registered with a wait that does not sleep. */
function setup(accounts: FakeAccount[] = [{ ...ACCOUNT, scopes: [...ACCOUNT.scopes] }]) {
  const graph = fakeGraph();
  const connections = fakeConnections({ accounts, provider: graph.provider, scopes: SCOPES });
  const fake = fakeApi({ connections: connections.connections, folder });
  fake.api.tools?.register(
    outlookTools(() => fake.api, {
      wait: async (ms) => {
        waits.push(ms);
      },
    }),
  );
  return { graph, connections, fake };
}

test("activate registers the read, write, generic and docs tools", async () => {
  const fake = fakeApi();
  stop = (await module.activate(fake.api)) ?? undefined;
  expect(fake.tools().map((tool) => tool.name)).toEqual([
    "outlook_search",
    "outlook_read",
    "outlook_folders",
    "outlook_draft_reply",
    "outlook_send",
    "outlook_reply",
    "outlook_reply_all",
    "outlook_forward",
    "outlook_move",
    "outlook_delete",
    "outlook_categories",
    "outlook_flag",
    "outlook_mark_read",
    "outlook_api",
    "outlook_api_docs",
  ]);
  expect(fake.ui).toEqual([]);
});

test("outlook_search lists id | date | from | subject | attachments, newest first", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_search", { query: "notice", max: 5 });
  expect(answer).toContain("id | date | from | subject | attachments");
  expect(answer).toContain(
    "m4 | 2026-10-11 06:00 | Dana Scully <dana@example.com> | RE: Partner Center notice | attachments",
  );
  expect(answer.indexOf("m4 |")).toBeLessThan(answer.indexOf("m3 |"));
  expect(answer).not.toContain("m1 |");
  const sent = graph.calls.find((call) => call.includes("$search"));
  expect(sent).toContain("$search=%22notice%22");
  expect(sent).not.toContain("$filter");
  expect(sent).not.toContain("$orderby");
  expect(sent).toContain("$select=");
});

test("outlook_search with no query lists the newest inbox mail, and says when nothing matched", async () => {
  const { fake } = setup();
  const newest = await fake.callTool("outlook_search", {});
  expect(newest.indexOf("m4 |")).toBeLessThan(newest.indexOf("m1 |"));
  expect(await fake.callTool("outlook_search", { query: "no such thing" })).toBe(
    "No mail matched.",
  );
});

test("outlook_read reads one message with its recipients and link", async () => {
  const { fake } = setup();
  const answer = await fake.callTool("outlook_read", { id: "m2" });
  expect(answer).toContain("From: Ada Lovelace <ada@example.com>");
  expect(answer).toContain("Cc: Dana Scully <dana@example.com>");
  expect(answer).toContain("Subject: RE: Contract for review");
  expect(answer).toContain("Looks fine to me, Bob.");
  expect(answer).toContain("Open in Outlook: https://outlook.office.com/mail/id/m2");
  expect(answer).not.toContain("Hi Ada");
});

test("outlook_read of a conversation reads it oldest first and saves every attachment", async () => {
  const { fake, graph } = setup();
  const answer = await fake.callTool("outlook_read", { id: "m1", conversation: true });
  expect(answer.indexOf("Hi Ada")).toBeLessThan(answer.indexOf("Looks fine"));
  expect(graph.calls.find((call) => call.includes("conversationId"))).not.toContain("$orderby");

  const png = /logo\.png[^\n]*saved to (\S+)/.exec(answer)?.[1];
  const pdf = /contract\.pdf[^\n]*saved to (\S+)/.exec(answer)?.[1];
  expect(png).toStartWith(join(folder, "mails"));
  expect([...(await readFile(png ?? ""))].slice(0, 4)).toEqual([0x89, 0x50, 0x4e, 0x47]);
  expect(await readFile(pdf ?? "", "utf8")).toBe("%PDF-1.4 fake contract");
});

test("a forwarded mail is saved as .eml and a reference attachment stays a link", async () => {
  const { fake } = setup();
  const forwarded = await fake.callTool("outlook_read", { id: "m3" });
  const eml = /Original notice[^\n]*saved to (.+\.eml)/.exec(forwarded)?.[1];
  expect(await readFile(eml ?? "", "utf8")).toContain("Subject: Original notice");

  const linked = await fake.callTool("outlook_read", { id: "m4" });
  expect(linked).toContain(
    "Notice.docx (text/html, 1 KB): link https://onedrive.example.com/notice.docx",
  );
});

test("attachments go under ctx.scratchpad/outlook/<hash>/ when the host gives one", async () => {
  const { fake } = setup();
  const tool = fake.tools().find((each) => each.name === "outlook_read");
  const scratchpad = join(folder, "scratch");
  const answer = await tool?.call({ id: "m1" }, {
    session: "s",
    pick: async () => null,
    signal: new AbortController().signal,
    scratchpad,
  } as never);
  expect(answer).toContain(join(scratchpad, "outlook"));
});

test("outlook_draft_reply asks for Mail.ReadWrite on the first draft, then saves a threaded draft", async () => {
  const { fake, connections } = setup();
  const answer = await fake.callTool("outlook_draft_reply", { id: "m1", body: "Thanks, signed." });
  expect(connections.accounts[0]?.scopes).toContain("Mail.ReadWrite");
  expect(answer).toContain('Reply draft saved in Drafts: "RE: Contract for review"');
  expect(answer).toContain("Draft id: draft1");
  expect(answer).toContain("Nothing was sent");
  expect(connections.requests.map((r) => `${r.method} ${new URL(r.url).pathname}`)).toContain(
    "POST /v1.0/me/messages/m1/createReply",
  );
});

test("reply-all uses createReplyAll, and replace_draft_id updates the draft's body", async () => {
  const { fake, graph } = setup([{ ...ACCOUNT, scopes: ["Mail.Read", "Mail.ReadWrite"] }]);
  await fake.callTool("outlook_draft_reply", { id: "m1", body: "First", all: true });
  expect(graph.calls).toContain("POST /me/messages/m1/createReplyAll");
  const answer = await fake.callTool("outlook_draft_reply", {
    id: "m1",
    body: "Second",
    replace_draft_id: "draft1",
  });
  expect(answer).toContain("Draft updated in Drafts");
  expect(graph.calls).toContain("PATCH /me/messages/draft1");
  expect(JSON.stringify(graph.messages.find((m) => m["id"] === "draft1")?.["body"])).toContain(
    "Second",
  );
});

test("a message that is not a draft is never overwritten", async () => {
  const { fake, graph } = setup([{ ...ACCOUNT, scopes: ["Mail.Read", "Mail.ReadWrite"] }]);
  const answer = await fake.callTool("outlook_draft_reply", {
    id: "m1",
    body: "x",
    replace_draft_id: "m2",
  });
  expect(answer).toBe("That id is not a draft; nothing was changed.");
  expect(graph.calls.some((call) => call.startsWith("PATCH"))).toBe(false);
});

test("a connection without Mail.Read is walked through incremental consent", async () => {
  const { fake, connections } = setup([{ ...ACCOUNT, scopes: ["User.Read"] }]);
  const asked: string[][] = [];
  const signIn = connections.connections.signIn;
  connections.connections.signIn = (provider, scopes, options) => {
    asked.push(scopes);
    return signIn(provider, scopes, options);
  };
  const answer = await fake.callTool("outlook_search", { query: "contract" });
  expect(asked).toEqual([["Mail.ReadWrite"]]);
  expect(answer).toContain("m1 |");
});

test("when the reader does not approve Mail.Read, the answer says to approve it in the browser", async () => {
  const { fake, connections } = setup([{ ...ACCOUNT, scopes: ["User.Read"] }]);
  connections.connections.signIn = async () => ({ ok: false, message: "Sign-in timed out." });
  const answer = await fake.callTool("outlook_search", { query: "x" });
  expect(answer).toContain(
    "Mail.ReadWrite was not granted for ada@example.com: Sign-in timed out.",
  );
  expect(answer).toContain("Approve prifly in the browser");
});

test("a consent refusal names the tenant admin and carries Microsoft's message", async () => {
  for (const code of ["AADSTS65001", "AADSTS90094", "AADSTS90095"]) {
    const { fake, connections } = setup([{ ...ACCOUNT, scopes: ["User.Read"] }]);
    const message = `${code}: Admin consent is required for this application.`;
    connections.connections.signIn = async () => ({ ok: false, message });
    const answer = await fake.callTool("outlook_search", { query: "x" });
    expect(answer).toContain("Ask your tenant admin to approve prifly");
    expect(answer).toContain(message);
    expect(answer).not.toContain("http");
  }
});

const ADMIN_URL = "https://login.microsoftonline.com/contoso/adminconsent?client_id=abc";

test("a consent refusal with an admin consent link puts the link in the answer", async () => {
  const { fake, connections } = setup([{ ...ACCOUNT, scopes: ["User.Read"] }]);
  const message = "AADSTS90094: Admin consent is required for this application.";
  connections.connections.signIn = async () => ({
    ok: false,
    message,
    adminConsentUrl: ADMIN_URL,
  });
  const answer = await fake.callTool("outlook_search", { query: "x" });
  expect(answer).toContain(`Ask your tenant admin to approve prifly: ${ADMIN_URL}`);
  expect(answer).toContain(`Microsoft said: ${message}`);
});

test("a needsSignIn answer with an admin consent link puts the link in the answer", async () => {
  const { fake, connections } = setup();
  connections.connections.fetch = async () => ({
    needsSignIn: {
      connection: ACCOUNT.id,
      provider: "microsoft",
      account: ACCOUNT.account,
      adminConsentUrl: ADMIN_URL,
    },
  });
  const answer = await fake.callTool("outlook_search", {});
  expect(answer).toContain("The Outlook grant for ada@example.com ended");
  expect(answer).toContain(`Ask your tenant admin to approve prifly: ${ADMIN_URL}`);
});

test("no connection, an unknown account and a missing host api are all plain text", async () => {
  const none = setup([]);
  expect(await none.fake.callTool("outlook_search", {})).toContain("Connect");
  expect(await none.fake.callTool("outlook_search", {})).toContain("Vault → Connections");

  const { fake } = setup();
  expect(await fake.callTool("outlook_search", { account: "nobody@example.com" })).toBe(
    'No connected Microsoft account matches "nobody@example.com". Connected: ada@example.com.',
  );

  const bare = fakeApi();
  bare.api.tools?.register(outlookTools(() => bare.api));
  expect(await bare.callTool("outlook_search", {})).toContain("update prifly");
});

test("the account argument and the account setting pick between connections", async () => {
  const second: FakeAccount = {
    id: "ms-2",
    provider: "microsoft",
    account: "bo@example.com",
    scopes: ["Mail.Read"],
  };
  const { fake, connections } = setup([{ ...ACCOUNT, scopes: ["Mail.Read"] }, second]);
  await fake.callTool("outlook_search", { account: "bo@example.com" });
  expect(connections.requests.at(-1)?.headers.get("Authorization")).toBe("Bearer token-ms-2-1");
  await fake.callTool("outlook_search", {});
  expect(connections.requests.at(-1)?.headers.get("Authorization")).toBe("Bearer token-ms-1-1");
});

test("a revoked grant comes back as the needs-sign-in text, not a throw", async () => {
  const { fake, connections } = setup();
  connections.endpoint.revoke("ms-1");
  const answer = await fake.callTool("outlook_search", {});
  expect(answer).toBe(
    "The Outlook grant for ada@example.com ended: sign in again in Vault → Connections.",
  );
});

test("a 401 claims challenge that reaches the extension is the same needs-sign-in text", async () => {
  const { fake, graph } = setup();
  graph.fault = { kind: "claims" };
  expect(await fake.callTool("outlook_search", {})).toBe(
    "The Outlook grant for ada@example.com ended: sign in again in Vault → Connections.",
  );
});

test("a 429 waits out Retry-After once and then retries", async () => {
  const { fake, graph } = setup();
  graph.fault = { kind: "throttle", status: 429, retryAfter: "2", times: 1 };
  const answer = await fake.callTool("outlook_search", {});
  expect(waits).toEqual([2000]);
  expect(answer).toContain("m4 |");
});

test("a second 429 gives text, after one wait capped at 30 s", async () => {
  const { fake, graph } = setup();
  graph.fault = { kind: "throttle", status: 429, retryAfter: "120", times: 5 };
  const answer = await fake.callTool("outlook_search", {});
  expect(waits).toEqual([30_000]);
  expect(answer).toBe("Microsoft Graph is throttling requests (429); try again in 120 s.");
});

test("a 503 without Retry-After is also waited out once", async () => {
  const { fake, graph } = setup();
  graph.fault = { kind: "throttle", status: 503, retryAfter: null, times: 5 };
  expect(await fake.callTool("outlook_search", {})).toContain("throttling requests (503)");
  expect(waits).toEqual([5000]);
});

test("a missing message and a missing argument are text", async () => {
  const { fake } = setup();
  expect(await fake.callTool("outlook_read", { id: "nope" })).toContain("ErrorItemNotFound");
  expect(await fake.callTool("outlook_read", {})).toContain("Give the id");
  expect(await fake.callTool("outlook_draft_reply", { id: "m1", body: " " })).toContain(
    "Give the id",
  );
});

test("searching and reading send only GET requests", async () => {
  const { fake, graph } = setup();
  await fake.callTool("outlook_search", { query: "contract" });
  await fake.callTool("outlook_read", { id: "m1", conversation: true });
  await fake.callTool("outlook_folders", {});
  expect(graph.calls.length).toBeGreaterThan(5);
  expect(graph.calls.filter((call) => !call.startsWith("GET "))).toEqual([]);
  expect(fake.confirms).toEqual([]);
});

test("a 503 on creating a draft is not repeated, so no second draft appears", async () => {
  const { fake, graph } = setup([{ ...ACCOUNT, scopes: ["Mail.Read", "Mail.ReadWrite"] }]);
  // The fault starts once the reader confirms: the GET of the original before it goes through.
  fake.api.confirm = async () => {
    graph.fault = { kind: "throttle", status: 503, retryAfter: "1", times: 5 };
    return true;
  };
  const answer = await fake.callTool("outlook_draft_reply", { id: "m1", body: "x" });
  expect(answer).toContain("throttling requests (503)");
  expect(waits).toEqual([]);
  expect(graph.calls.filter((call) => !call.startsWith("GET "))).toEqual([
    "POST /me/messages/m1/createReply",
  ]);
});
