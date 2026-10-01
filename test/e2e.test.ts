import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { createHash, randomBytes } from "node:crypto";
import http, { type Server } from "node:http";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createApp } from "../src/server.js";
import { MemoryStore, PostgresStore } from "../src/store.js";
import type { Config } from "../src/config.js";
import { startFakeMicrosoft, PNG } from "./fake-microsoft.js";
import { existsSync } from "node:fs";
import { closeBrowser } from "../src/diagrams.js";

const LOCAL_CHROMIUM = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
if (!process.env.CHROMIUM_PATH && existsSync(LOCAL_CHROMIUM)) process.env.CHROMIUM_PATH = LOCAL_CHROMIUM;

let fake: Awaited<ReturnType<typeof startFakeMicrosoft>>;
let server: Server;
let base: string;
const store = process.env.TEST_DATABASE_URL ? new PostgresStore(process.env.TEST_DATABASE_URL) : new MemoryStore();
const REDIRECT = "http://localhost:9999/callback";

before(async () => {
  fake = await startFakeMicrosoft();
  // Bind first to learn the port, then build the app with the real base URL.
  let handler: http.RequestListener = () => {};
  server = http.createServer((req, res) => handler(req, res));
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  base = `http://127.0.0.1:${(server.address() as any).port}`;
  handler = createApp(makeConfig(base), store).app;
});

after(async () => {
  await closeBrowser();
  server?.close();
  if (store instanceof PostgresStore) await (store as any).pool.end();
  fake?.server.close();
});

function makeConfig(baseUrl: string): Config {
  return {
    port: 0,
    baseUrl: new URL(baseUrl),
    msClientId: "client",
    msClientSecret: "secret",
    msTenant: "common",
    msAuthority: fake.url,
    graphBase: `${fake.url}/v1.0`,
    msScopes: ["offline_access", "User.Read", "Notes.ReadWrite"],
    encryptionKey: randomBytes(32),
  };
}

const pkce = () => {
  const verifier = randomBytes(32).toString("base64url");
  return { verifier, challenge: createHash("sha256").update(verifier).digest("base64url") };
};

async function register() {
  const r = await fetch(`${base}/register`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ client_name: "Claude <script>", redirect_uris: [REDIRECT], token_endpoint_auth_method: "none" }),
  });
  assert.equal(r.status, 201);
  return (await r.json()) as { client_id: string };
}

/** Runs the full browser flow and returns our tokens. */
async function signIn(clientId: string) {
  const { verifier, challenge } = pkce();
  const authUrl = new URL(`${base}/authorize`);
  authUrl.search = new URLSearchParams({
    response_type: "code", client_id: clientId, redirect_uri: REDIRECT, code_challenge: challenge,
    code_challenge_method: "S256", state: "xyz", resource: `${base}/mcp`,
  }).toString();
  const consent = await fetch(authUrl, { redirect: "manual" });
  assert.equal(consent.status, 200);
  const html = await consent.text();
  assert.match(html, /Claude &lt;script&gt; wants to access your OneNote/, "client name is escaped");
  assert.equal(consent.headers.get("x-frame-options"), "DENY");
  const pendingId = html.match(/name="pending_id" value="([^"]+)"/)![1];
  const csrf = html.match(/name="csrf" value="([^"]+)"/)![1];
  const cookie = consent.headers.get("set-cookie")!.split(";")[0];

  const allow = await fetch(`${base}/oauth/consent`, {
    method: "POST", redirect: "manual",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Cookie: cookie },
    body: new URLSearchParams({ pending_id: pendingId, csrf, action: "allow" }),
  });
  assert.equal(allow.status, 302);
  const msUrl = allow.headers.get("location")!;
  assert.ok(msUrl.startsWith(`${fake.url}/common/oauth2/v2.0/authorize`));
  assert.match(msUrl, /scope=offline_access/);

  const msRedirect = await fetch(msUrl, { redirect: "manual" });
  const callback = await fetch(msRedirect.headers.get("location")!, { redirect: "manual" });
  assert.equal(callback.status, 302);
  const back = new URL(callback.headers.get("location")!);
  assert.equal(back.origin + back.pathname, REDIRECT);
  assert.equal(back.searchParams.get("state"), "xyz");
  const code = back.searchParams.get("code")!;

  const tok = await fetch(`${base}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "authorization_code", code, code_verifier: verifier, client_id: clientId, redirect_uri: REDIRECT }),
  });
  assert.equal(tok.status, 200, await tok.clone().text());
  return (await tok.json()) as { access_token: string; refresh_token: string; expires_in: number };
}

async function mcp(accessToken: string) {
  const client = new Client({ name: "test", version: "1" });
  await client.connect(new StreamableHTTPClientTransport(new URL(`${base}/mcp`), { requestInit: { headers: { Authorization: `Bearer ${accessToken}` } } }));
  const call = async (name: string, args: Record<string, unknown> = {}) => {
    const r: any = await client.callTool({ name, arguments: args });
    return { ...r, text: r.content?.[0]?.text as string, data: (() => { try { return JSON.parse(r.content?.[0]?.text); } catch { return undefined; } })() };
  };
  return { client, call };
}

// ---------------------------------------------------------------------------

test("discovery metadata is advertised", async () => {
  const pr = await (await fetch(`${base}/.well-known/oauth-protected-resource/mcp`)).json();
  assert.equal(pr.resource, `${base}/mcp`);
  const as = await (await fetch(`${base}/.well-known/oauth-authorization-server`)).json();
  assert.equal(as.registration_endpoint, `${base}/register`);
  assert.deepEqual(as.code_challenge_methods_supported, ["S256"]);

  const unauth = await fetch(`${base}/mcp`, { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  assert.equal(unauth.status, 401);
  assert.match(unauth.headers.get("www-authenticate")!, /resource_metadata=/);
});

test("consent cannot be forged cross-site (no cookie) and cancel returns access_denied", async () => {
  const { client_id } = await register();
  const { challenge } = pkce();
  const q = new URLSearchParams({ response_type: "code", client_id, redirect_uri: REDIRECT, code_challenge: challenge, code_challenge_method: "S256", state: "s" });
  const page = await fetch(`${base}/authorize?${q}`);
  const html = await page.text();
  const pendingId = html.match(/name="pending_id" value="([^"]+)"/)![1];
  const csrf = html.match(/name="csrf" value="([^"]+)"/)![1];
  const forged = await fetch(`${base}/oauth/consent`, {
    method: "POST", redirect: "manual", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ pending_id: pendingId, csrf, action: "allow" }),
  });
  assert.equal(forged.status, 400);

  const page2 = await fetch(`${base}/authorize?${q}`);
  const html2 = await page2.text();
  const cancel = await fetch(`${base}/oauth/consent`, {
    method: "POST", redirect: "manual",
    headers: { "Content-Type": "application/x-www-form-urlencoded", Cookie: page2.headers.get("set-cookie")!.split(";")[0] },
    body: new URLSearchParams({ pending_id: html2.match(/name="pending_id" value="([^"]+)"/)![1], csrf: html2.match(/name="csrf" value="([^"]+)"/)![1], action: "deny" }),
  });
  const loc = new URL(cancel.headers.get("location")!);
  assert.equal(loc.searchParams.get("error"), "access_denied");
  assert.equal(loc.searchParams.get("state"), "s");
});

test("unregistered redirect_uri is rejected", async () => {
  const { client_id } = await register();
  const { challenge } = pkce();
  const r = await fetch(`${base}/authorize?${new URLSearchParams({ response_type: "code", client_id, redirect_uri: "https://evil.example/cb", code_challenge: challenge, code_challenge_method: "S256" })}`, { redirect: "manual" });
  assert.equal(r.status, 400);
});

test("invalid authorization code is rejected", async () => {
  const { client_id } = await register();
  // Use the real flow but swap the verifier.
  const tokens = await signIn(client_id);
  assert.ok(tokens.access_token);
  const bad = await fetch(`${base}/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "authorization_code", code: "nope", code_verifier: "x".repeat(43), client_id }),
  });
  assert.equal(bad.status, 400);
});

test("full OneNote workflow through MCP", async () => {
  const { client_id } = await register();
  const tokens = await signIn(client_id);
  const { client, call } = await mcp(tokens.access_token);

  const tools = (await client.listTools()).tools.map((t) => t.name).sort();
  assert.deepEqual(tools, [
    "append_to_page", "copy_section", "create_notebook", "create_page", "create_section", "create_section_group",
    "delete_page", "get_notebook_structure", "get_page_resource", "insert_diagram", "list_notebooks", "list_pages", "move_page",
    "read_page", "search_pages", "update_page", "whoami",
  ]);
  const del = (await client.listTools()).tools.find((t) => t.name === "delete_page")!;
  assert.equal(del.annotations?.destructiveHint, true);

  assert.equal((await call("whoami")).data.name, "Test Student");
  assert.equal((await call("list_notebooks")).data[0].notebook_id, "nb1");

  const tree = (await call("get_notebook_structure")).text;
  assert.match(tree, /📓 Uni/);
  assert.match(tree, /📁 Year 1 {2}\[section_group_id: sg1\]\n {4}📄 Maths {2}\[section_id: s2\]/);
  assert.match(tree, /📄 Inbox/);

  // Read with markdown conversion, to-dos and image references.
  const page = (await call("read_page", { page_id: "p1" })).text;
  assert.match(page, /^# Lecture 1 - Vectors/);
  assert.match(page, /# Vectors/);
  assert.match(page, /- \[ \] Do problem sheet/);
  assert.match(page, /!\[diagram\]\(onenote-resource:r1\)/);

  const img: any = await client.callTool({ name: "get_page_resource", arguments: { resource_id: "r1" } });
  assert.equal(img.content[0].type, "image");
  assert.equal(img.content[0].data, PNG.toString("base64"));

  // Create a page from Markdown.
  const created = (await call("create_page", {
    section_id: "s2", title: "Lecture 2 - Matrices",
    content: "## Key ideas\n\n- Rows & columns\n\n- [ ] Revise determinants\n- [x] Watch lecture\n\n| a | b |\n|---|---|\n| 1 | 2 |",
  })).data;
  const html = fake.pages[created.page_id].body;
  assert.match(html, /<h2[^>]*>Key ideas<\/h2>/);
  assert.match(html, /<p data-tag="to-do">Revise determinants<\/p>/);
  assert.match(html, /<p data-tag="to-do:completed">Watch lecture<\/p>/);
  assert.match(html, /<table border="1">/);
  assert.equal(fake.pages[created.page_id].title, "Lecture 2 - Matrices");

  await call("append_to_page", { page_id: created.page_id, content: "**Summary:** done" });
  assert.match(fake.pages[created.page_id].body, /<strong>Summary:<\/strong> done/);

  // Targeted edit + rename.
  const rawHtml = (await call("read_page", { page_id: "p1", format: "html" })).text;
  assert.match(rawHtml, /data-id="intro"/);
  await call("update_page", { page_id: "p1", title: "L1 Vectors", edits: [{ target: "#intro", action: "replace", content: "Vectors: size + direction." }] });
  assert.equal(fake.pages.p1.title, "L1 Vectors");
  assert.match(fake.pages.p1.body, /Vectors: size \+ direction\./);
  assert.ok(!fake.pages.p1.body.includes("magnitude"));

  // Paging across nextLink (fake returns one item per page).
  const listed = (await call("list_pages", { section_id: "s2" })).data;
  assert.equal(listed.length, 2);

  // Search by title and content.
  const s1 = (await call("search_pages", { query: "matrices" })).data;
  assert.equal(s1.title_matches[0].title, "Lecture 2 - Matrices");
  const s2 = (await call("search_pages", { query: "eigenvalues", notebook_id: "nb1", search_content: true })).data;
  assert.equal(s2.content_matches[0].page_id, "p2");
  assert.match(s2.content_matches[0].snippet, /Eigenvalues are cool/);
  const s3 = await call("search_pages", { query: "x", search_content: true });
  assert.equal(s3.isError, true);

  // Organise: new section group + section, then move a page.
  const sg = (await call("create_section_group", { name: "Year 2", notebook_id: "nb1" })).data;
  const sec = (await call("create_section", { name: "Physics", section_group_id: sg.section_group_id })).data;
  const bad = await call("create_section", { name: "x", notebook_id: "nb1", section_group_id: "sg1" });
  assert.equal(bad.isError, true);
  const moved = (await call("move_page", { page_id: "p2", target_section_id: sec.section_id })).data;
  assert.equal(moved.moved, true);
  assert.equal(fake.pages.p2, undefined, "original deleted");
  assert.equal(fake.pages[moved.new_page].section, sec.section_id);

  const copied = (await call("copy_section", { section_id: "s1", target_notebook_id: "nb1", new_name: "Inbox copy" })).data;
  assert.equal(fake.sections[copied.new_section].displayName, "Inbox copy");

  await call("delete_page", { page_id: created.page_id });
  assert.equal(fake.pages[created.page_id], undefined);

  // Graph errors come back as readable tool errors.
  const missing = await call("read_page", { page_id: "nope" });
  assert.equal(missing.isError, true);
  assert.match(missing.text, /404/);

  await client.close();
});

test("throttling is retried transparently", async () => {
  const { client_id } = await register();
  const { access_token } = await signIn(client_id);
  const { call, client } = await mcp(access_token);
  fake.state.throttleNext = 2;
  assert.equal((await call("list_notebooks")).data[0].name, "Uni");
  await client.close();
});

test("expired Microsoft token is refreshed; revoked sign-in forces reconnect", async () => {
  const { client_id } = await register();
  const tokens = await signIn(client_id);
  const { call, client } = await mcp(tokens.access_token);

  // Microsoft invalidates the access token → we get 401 from Graph → refresh → retry succeeds.
  fake.state.validAccess.clear();
  const before = fake.state.refreshCount;
  assert.equal((await call("whoami")).data.name, "Test Student");
  assert.equal(fake.state.refreshCount, before + 1);

  // Our own refresh_token rotates.
  const r = await fetch(`${base}/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: tokens.refresh_token, client_id }),
  });
  assert.equal(r.status, 200);
  const reuse = await fetch(`${base}/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: tokens.refresh_token, client_id }),
  });
  assert.equal(reuse.status, 400, "old refresh token cannot be reused");

  // User revokes access at Microsoft → tool says reconnect, and the next request is 401 so Claude re-auths.
  fake.state.validAccess.clear();
  fake.state.validRefresh.clear();
  const res = await call("whoami");
  assert.equal(res.isError, true);
  assert.match(res.text, /Reconnect/);
  await client.close();
  const after = await fetch(`${base}/mcp`, {
    method: "POST",
    headers: { Authorization: `Bearer ${tokens.access_token}`, "Content-Type": "application/json", Accept: "application/json, text/event-stream" },
    body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list" }),
  });
  assert.equal(after.status, 401);
});

test("Microsoft tokens are encrypted at rest", async () => {
  const dump = store instanceof MemoryStore
    ? JSON.stringify([...(store as any).data.entries()])
    : JSON.stringify((await (store as any).pool.query("SELECT * FROM kv")).rows);
  assert.ok(!/msrt\d+/.test(dump), "no plaintext Microsoft refresh tokens in storage");
  assert.ok(!/msat\d+/.test(dump), "no plaintext Microsoft access tokens in storage");
});

const isPng = (b: Buffer) => b.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));

test("diagrams: Mermaid and SVG blocks become images on the page", async () => {
  const { client_id } = await register();
  const { access_token } = await signIn(client_id);
  const { client, call } = await mcp(access_token);

  const created = await call("create_page", {
    section_id: "s2",
    title: "Control systems",
    content: [
      "## Loop",
      "```mermaid Closed-loop motor control",
      "flowchart LR",
      "  Setpoint --> Controller --> Motor --> Sensor --> Controller",
      "```",
      "## Free-body diagram",
      "```svg",
      '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><rect x="70" y="70" width="60" height="60" fill="#ddd" stroke="#000"/>' +
        '<line x1="100" y1="130" x2="100" y2="190" stroke="red" stroke-width="3"/><text x="105" y="180" font-size="14">mg</text></svg>',
      "```",
      "```python",
      "print('normal code stays code')",
      "```",
    ].join("\n"),
  });
  assert.equal(created.isError, undefined, created.text);
  assert.equal(created.data.diagrams, 2);

  const upload = fake.multipartLog.at(-1)!.parts;
  assert.ok(upload.Presentation, "multipart has Presentation part");
  assert.ok(isPng(upload.diagram1.data) && upload.diagram1.type === "image/png");
  assert.ok(isPng(upload.diagram2.data));
  const presentation = upload.Presentation.data.toString();
  assert.match(presentation, /<img src="name:diagram1" width="\d+" alt="Closed-loop motor control" \/>/);
  assert.match(presentation, /<img src="name:diagram2" width="200" alt="svg diagram" \/>/);
  assert.match(presentation, /print\(&#39;normal code stays code&#39;\)/);

  // Reading it back shows image references that can be fetched.
  const page = (await call("read_page", { page_id: created.data.page_id })).text;
  const refs = [...page.matchAll(/onenote-resource:([\w]+)/g)].map((m) => m[1]);
  assert.equal(refs.length, 2);
  const img: any = await client.callTool({ name: "get_page_resource", arguments: { resource_id: refs[0] } });
  assert.equal(img.content[0].type, "image");

  // insert_diagram appends to an existing page via multipart PATCH.
  const added = await call("insert_diagram", { page_id: "p1", kind: "mermaid", source: "stateDiagram-v2\n  [*] --> Idle\n  Idle --> Run", caption: "Robot states" });
  assert.equal(added.isError, undefined, added.text);
  const patch = fake.multipartLog.at(-1)!.parts;
  assert.ok(patch.Commands && isPng(patch.diagram1.data));
  assert.match(fake.pages.p1.body, /onenote\/resources\/res\d+/);
  assert.match(fake.pages.p1.body, /<em>Robot states<\/em>/);

  // insert_diagram can also create a new page.
  const fresh = await call("insert_diagram", { section_id: "s1", kind: "svg", source: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4"/></svg>', title: "Circle" });
  assert.equal(fake.pages[fresh.data.page_id].title, "Circle");

  // update_page edits can carry diagrams too.
  const upd = await call("update_page", { page_id: "p1", edits: [{ target: "#intro", action: "replace", content: "```svg\n<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20\" height=\"20\"><rect width=\"20\" height=\"20\"/></svg>\n```" }] });
  assert.equal(upd.isError, undefined, upd.text);
  assert.ok(fake.multipartLog.at(-1)!.parts.edit1diagram1);

  // Errors are reported so Claude can fix the source; nothing is written.
  const before = Object.keys(fake.pages).length;
  const badMermaid = await call("create_page", { section_id: "s1", title: "x", content: "```mermaid\nflowchart LR\n  A -->\n```" });
  assert.equal(badMermaid.isError, true);
  assert.match(badMermaid.text, /Diagram 1 \(mermaid\) failed to render\. Mermaid syntax error/);
  const badSvg = await call("insert_diagram", { section_id: "s1", kind: "svg", source: "<svg><rect" });
  assert.equal(badSvg.isError, true);
  assert.match(badSvg.text, /Invalid SVG/);
  const notSvg = await call("insert_diagram", { section_id: "s1", kind: "svg", source: "hello" });
  assert.match(notSvg.text, /must contain an <svg> element/);
  assert.equal(Object.keys(fake.pages).length, before);

  await client.close();
});

test("diagrams: SVG can't reach the network or run scripts; Mermaid page has no network", async () => {
  const { renderSvg, renderMermaid } = await import("../src/diagrams.js");
  // External image href + script are ignored rather than fetched/executed.
  const r = renderSvg('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="50" height="50"><script>alert(1)</script><image xlink:href="http://127.0.0.1:1/x.png" width="50" height="50"/></svg>');
  assert.ok(isPng(r.png));
  // Mermaid click directives with javascript: links are blocked by securityLevel strict and still render.
  const m = await renderMermaid('flowchart LR\n  A --> B\n  click A "javascript:alert(1)"');
  assert.ok(isPng(m.png));
});
