import { expect, test } from "bun:test";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import { boot, cleanUp } from "./support";

cleanUp();

/** Everything the board, the launch, the open and the poller do, once. */
async function useEverything() {
  const ext = await boot();
  await ext.board();
  await ext.board("contract");
  await ext.board("from:dana");
  await ext.extension.open("mail", "convA");
  await ext.extension.open("mail", "convB");
  await ext.extension.launch("mail", "convA", "", { again: false });
  await ext.extension.launch("mail", "convB", "", { again: false });
  await ext.extension.launched("mail", "convB", "session-1");
  await ext.extension.launch("mail", "row:grant", "", { again: false });
  await ext.extension.launch("mail", "action:switch", "ada@example.com", { again: false });
  ext.graph.messages.push({
    ...ext.graph.messages[3],
    id: "m9",
    receivedDateTime: "2026-10-11T11:00:00Z",
  });
  await ext.poll();
  ext.graph.expiredDelta = true;
  await ext.poll();
  return ext;
}

test("the board, launch, open and poller send only GET requests, to no send path", async () => {
  const ext = await useEverything();
  const requests = ext.connections.requests;
  expect(requests.length).toBeGreaterThan(30);
  for (const request of requests) {
    expect(request.method).toBe("GET");
    expect(new URL(request.url).pathname.toLowerCase()).not.toMatch(
      /\/(send|sendmail|reply|replyall|forward|createreply|createreplyall|createforward)$/,
    );
  }
  expect(ext.graph.calls.filter((call) => !call.startsWith("GET "))).toEqual([]);
});

test("the Graph client and everything built on it have no call that sends", async () => {
  const files = (await readdir(join(import.meta.dir, ".."))).filter((name) => name.endsWith(".ts"));
  expect(files).toContain("launch.ts");
  expect(files).toContain("poll.ts");
  for (const file of files) {
    const source = await readFile(join(import.meta.dir, "..", file), "utf8");
    // No sendMail, no /send, no /forward, no direct /reply (createReply makes a draft and is graph.ts's alone).
    expect(source, file).not.toMatch(
      /sendMail|["'`/]send["'`/]|\/forward\b|\/reply(All)?\b|createForward/i,
    );
    if (file !== "graph.ts" && file !== "tools.ts") {
      expect(source, file).not.toMatch(
        /createReply|updateDraft|method:\s*["'](POST|PATCH|PUT|DELETE)/,
      );
    }
  }
});
