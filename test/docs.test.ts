import { expect, test } from "bun:test";
import { pageFor, referenceText } from "../docs";
import { setup } from "./harness";

/** A Learn reference page as Microsoft serves it, cut down to its shape. */
const MOVE_PAGE = `<html><body><main>
<h1 id="message-move">message: move</h1>
<p>Move a message to another folder.</p>
<h2 id="permissions">Permissions</h2>
<table><thead><tr><th>Permission type</th><th>Least privileged permissions</th></tr></thead>
<tbody><tr><td>Delegated (work or school account)</td><td>Mail.ReadWrite</td></tr></tbody></table>
<h2 id="http-request">HTTP request</h2>
<!-- { "blockType": "ignored" } -->
<pre><code class="lang-http">POST /me/messages/{id}/move
</code></pre>
<h2 id="request-body">Request body</h2>
<table><tbody><tr><td>destinationId</td><td>String</td><td>The destination folder ID, or a well-known folder name.</td></tr></tbody></table>
<h2 id="response">Response</h2>
<p>If successful, this method returns <code>201 Created</code> &amp; a message.</p>
<h2 id="example">Example</h2>
<div class="tabGroup"><pre><code class="lang-http">POST https://graph.microsoft.com/v1.0/me/messages/AAMkADhAAATs28OAAA=/move
Content-type: application/json

{ &quot;destinationId&quot;: &quot;deleteditems&quot; }
</code></pre>
<pre><code class="lang-csharp">var x = new MovePostRequestBody();</code></pre></div>
<h2 id="ms--feedback">Feedback</h2>
</main></body></html>`;

test("a topic or a Graph path becomes a reference page name", () => {
  const cases: [string, string][] = [
    ["sendMail", "api/user-sendmail"],
    ["create event", "api/user-post-events"],
    ["message-move", "api/message-move"],
    ["POST /me/messages/{id}/move", "api/message-move"],
    ["/me/sendMail", "api/user-sendmail"],
    ["GET /me/events", "api/user-list-events"],
    ["POST /me/events", "api/user-post-events"],
    ["PATCH /me/events/{id}", "api/event-update"],
    ["DELETE /v1.0/me/messages/AAMkAD=", "api/message-delete"],
    ["/me/mailFolders/inbox/messages", "api/mailfolder-list-messages"],
    ["/me/calendarView", "api/user-list-calendarview"],
    ["/me/messages/delta", "api/message-delta"],
    [
      "POST /me/messages/{id}/attachments/createUploadSession",
      "api/attachment-createuploadsession",
    ],
    ["/me/outlook/masterCategories", "api/outlookuser-list-mastercategories"],
    ["POST /me/events/{id}/accept", "api/event-accept"],
    ["https://graph.microsoft.com/v1.0/me/messages", "api/user-list-messages"],
    [
      "https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0",
      "api/user-sendmail",
    ],
  ];
  for (const [topic, page] of cases) expect(pageFor(topic), topic).toBe(page);
  expect(pageFor("something nobody maps")).toBeNull();
});

test("the reference sections come back as text, the example request with them, the rest left out", () => {
  const text = referenceText(MOVE_PAGE);
  expect(text).toStartWith("# message: move");
  expect(text).toContain("Mail.ReadWrite");
  expect(text).toContain("```\nPOST /me/messages/{id}/move\n```");
  expect(text).toContain("| destinationId | String | The destination folder ID");
  expect(text).toContain("201 Created & a message");
  expect(text).toContain("## Example request");
  expect(text).toContain('{ "destinationId": "deleteditems" }');
  expect(text).not.toContain("MovePostRequestBody");
  expect(text).not.toContain("Feedback");
  expect(text).not.toContain("blockType");
});

test("outlook_api_docs fetches the page from Microsoft Learn, falling back to Learn's search", async () => {
  const fetched: string[] = [];
  const fetchDocs = async (url: string) => {
    fetched.push(url);
    if (url.startsWith("https://learn.microsoft.com/api/search")) {
      return Response.json({
        results: [
          { url: "https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0" },
          { url: "https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0" },
        ],
      });
    }
    if (url.includes("/graph/api/message-move?view=graph-rest-1.0")) return new Response(MOVE_PAGE);
    return new Response("not found", { status: 404 });
  };
  const { fake, connections } = setup({ deps: { fetchDocs } });

  const direct = await fake.callTool("outlook_api_docs", { topic: "POST /me/messages/{id}/move" });
  expect(direct).toStartWith(
    "From https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0",
  );
  expect(direct).toContain("destinationId");

  const searched = await fake.callTool("outlook_api_docs", { topic: "relocate a mail" });
  expect(searched).toContain("destinationId");
  expect(fetched.some((url) => url.includes("search=relocate+a+mail"))).toBe(true);

  for (const url of fetched) expect(url).toStartWith("https://learn.microsoft.com/");
  expect(connections.requests).toEqual([]);
  expect(fake.confirms).toEqual([]);
});

test("a topic with no page says so", async () => {
  const fetchDocs = async (url: string) =>
    url.includes("/api/search")
      ? Response.json({ results: [] })
      : new Response("", { status: 404 });
  const { fake } = setup({ deps: { fetchDocs } });
  const answer = await fake.callTool("outlook_api_docs", { topic: "POST /me/nothing/{id}/zap" });
  expect(answer).toContain("No Microsoft Graph v1.0 reference page found");
  expect(answer).toContain("tried api/");
});
