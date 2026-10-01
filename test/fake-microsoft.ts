/**
 * A small fake of the Microsoft identity platform + Graph OneNote API, for end-to-end tests.
 */
import express from "express";
import type { Server } from "node:http";
import { createHash } from "node:crypto";

// 1x1 PNG
export const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64");

export function startFakeMicrosoft() {
  const app = express();
  app.use(express.urlencoded({ extended: false }));
  app.use(express.json());
  app.use(express.text({ type: ["text/html", "application/xhtml+xml"] }));

  const state = {
    codes: new Map<string, { challenge: string; redirect: string }>(),
    validAccess: new Set<string>(),
    validRefresh: new Set<string>(),
    refreshCount: 0,
    graphCalls: [] as string[],
    throttleNext: 0,
    seq: 0,
  };
  const id = (p: string) => `${p}${++state.seq}`;

  const nb = [{ id: "nb1", displayName: "Uni", lastModifiedDateTime: "2026-09-30T10:00:00Z", isShared: false, userRole: "Owner" }];
  const groups: Record<string, { id: string; displayName: string; parent: string }> = {
    sg1: { id: "sg1", displayName: "Year 1", parent: "nb1" },
  };
  const sections: Record<string, { id: string; displayName: string; parent: string }> = {
    s1: { id: "s1", displayName: "Inbox", parent: "nb1" },
    s2: { id: "s2", displayName: "Maths", parent: "sg1" },
  };
  const pages: Record<string, { id: string; title: string; section: string; body: string; order: number }> = {
    p1: {
      id: "p1", title: "Lecture 1 - Vectors", section: "s2", order: 0,
      body: `<h1 data-id="h">Vectors</h1><p data-id="intro">A vector has magnitude and direction.</p><img src="https://graph.microsoft.com/v1.0/users('x')/onenote/resources/r1/$value" alt="diagram"/><p data-tag="to-do">Do problem sheet</p>`,
    },
    p2: { id: "p2", title: "Random thoughts", section: "s1", order: 0, body: `<p>Eigenvalues are cool</p>` },
  };
  const ops: Record<string, { polls: number; resourceId: string }> = {};

  // ---------- identity ----------
  app.get("/common/oauth2/v2.0/authorize", (req, res) => {
    const code = id("mscode");
    state.codes.set(code, { challenge: String(req.query.code_challenge), redirect: String(req.query.redirect_uri) });
    const u = new URL(String(req.query.redirect_uri));
    u.searchParams.set("code", code);
    u.searchParams.set("state", String(req.query.state));
    res.redirect(302, u.toString());
  });

  const issue = (res: express.Response) => {
    const access = id("msat");
    const refresh = id("msrt");
    state.validAccess.add(access);
    state.validRefresh.add(refresh);
    res.json({ access_token: access, refresh_token: refresh, expires_in: 3600, token_type: "Bearer" });
  };

  app.post("/common/oauth2/v2.0/token", (req, res) => {
    if (req.body.client_secret !== "secret") return res.status(401).json({ error: "invalid_client" });
    if (req.body.grant_type === "authorization_code") {
      const c = state.codes.get(req.body.code);
      state.codes.delete(req.body.code);
      const ok = c && createHash("sha256").update(req.body.code_verifier).digest("base64url") === c.challenge && c.redirect === req.body.redirect_uri;
      if (!ok) return res.status(400).json({ error: "invalid_grant" });
      return issue(res);
    }
    if (req.body.grant_type === "refresh_token") {
      if (!state.validRefresh.has(req.body.refresh_token)) return res.status(400).json({ error: "invalid_grant", error_description: "revoked" });
      state.validRefresh.delete(req.body.refresh_token);
      state.refreshCount++;
      return issue(res);
    }
    res.status(400).json({ error: "unsupported_grant_type" });
  });

  // ---------- graph ----------
  const g = express.Router();
  g.use((req, res, next) => {
    state.graphCalls.push(`${req.method} ${req.path}`);
    const tok = (req.headers.authorization ?? "").replace("Bearer ", "");
    if (!state.validAccess.has(tok)) return res.status(401).json({ error: { code: "InvalidAuthenticationToken", message: "expired" } });
    if (state.throttleNext > 0) {
      state.throttleNext--;
      return res.status(429).set("Retry-After", "0").json({ error: { code: "20166", message: "throttled" } });
    }
    next();
  });
  const notFound = (res: express.Response) => res.status(404).json({ error: { code: "20102", message: "The requested resource does not exist." } });
  const pageJson = (p: (typeof pages)[string]) => ({
    id: p.id, title: p.title, order: p.order, level: 0, lastModifiedDateTime: "2026-09-30T10:00:00Z",
    parentSection: { id: p.section, displayName: sections[p.section]?.displayName },
    parentNotebook: { displayName: "Uni" },
    links: { oneNoteWebUrl: { href: `https://onenote.example/${p.id}` } },
  });
  // Paginate with page size 1 to exercise @odata.nextLink handling.
  const paged = (req: express.Request, res: express.Response, items: unknown[]) => {
    const skip = Number(req.query.$skip ?? 0);
    const out: any = { value: items.slice(skip, skip + 1) };
    if (skip + 1 < items.length) {
      const u = new URL(`http://${req.headers.host}${req.baseUrl}${req.path}`);
      for (const [k, v] of Object.entries(req.query)) u.searchParams.set(k, String(v));
      u.searchParams.set("$skip", String(skip + 1));
      out["@odata.nextLink"] = u.toString();
    }
    res.json(out);
  };

  g.get("/me", (_req, res) => res.json({ displayName: "Test Student", mail: "student@example.ac.uk" }));
  g.get("/me/onenote/notebooks", (req, res) => paged(req, res, nb));
  g.get("/me/onenote/notebooks/:id", (req, res) => { const n = nb.find((x) => x.id === req.params.id); n ? res.json(n) : notFound(res); });
  g.post("/me/onenote/notebooks", (req, res) => { const n = { id: id("nb"), displayName: req.body.displayName, lastModifiedDateTime: "", isShared: false, userRole: "Owner" }; nb.push(n); res.status(201).json(n); });
  for (const kind of ["notebooks", "sectionGroups"]) {
    g.get(`/me/onenote/${kind}/:id/sections`, (req, res) => paged(req, res, Object.values(sections).filter((s) => s.parent === req.params.id)));
    g.get(`/me/onenote/${kind}/:id/sectionGroups`, (req, res) => paged(req, res, Object.values(groups).filter((s) => s.parent === req.params.id)));
    g.post(`/me/onenote/${kind}/:id/sections`, (req, res) => { const s = { id: id("s"), displayName: req.body.displayName, parent: req.params.id }; sections[s.id] = s; res.status(201).json(s); });
    g.post(`/me/onenote/${kind}/:id/sectionGroups`, (req, res) => { const s = { id: id("sg"), displayName: req.body.displayName, parent: req.params.id }; groups[s.id] = s; res.status(201).json(s); });
  }
  g.get("/me/onenote/sections/:id/pages", (req, res) => {
    if (!sections[req.params.id]) return notFound(res);
    paged(req, res, Object.values(pages).filter((p) => p.section === req.params.id).map(pageJson));
  });
  g.get("/me/onenote/pages", (req, res) => {
    const m = String(req.query.$filter ?? "").match(/contains\(tolower\(title\),'((?:[^']|'')*)'\)/);
    const q = m ? m[1].replace(/''/g, "'") : "";
    res.json({ value: Object.values(pages).filter((p) => p.title.toLowerCase().includes(q)).map(pageJson) });
  });
  g.post("/me/onenote/sections/:id/pages", (req, res) => {
    const html = String(req.body);
    const title = html.match(/<title>(.*?)<\/title>/)?.[1] ?? "";
    const body = html.match(/<body>([\s\S]*)<\/body>/)?.[1] ?? "";
    const p = { id: id("p"), title, section: req.params.id, body, order: 99 };
    pages[p.id] = p;
    res.status(201).json(pageJson(p));
  });
  g.get("/me/onenote/pages/:id", (req, res) => { const p = pages[req.params.id]; p ? res.json(pageJson(p)) : notFound(res); });
  g.get("/me/onenote/pages/:id/content", (req, res) => {
    const p = pages[req.params.id];
    if (!p) return notFound(res);
    res.type("text/html").send(`<html><head><title>${p.title}</title></head><body>${p.body}</body></html>`);
  });
  g.patch("/me/onenote/pages/:id/content", (req, res) => {
    const p = pages[req.params.id];
    if (!p) return notFound(res);
    for (const c of req.body as any[]) {
      if (c.target === "title" && c.action === "replace") p.title = c.content;
      else if (c.target === "body" && c.action === "append") p.body += c.content;
      else if (c.target === "body" && c.action === "prepend") p.body = c.content + p.body;
      else if (c.target.startsWith("#") && c.action === "replace") {
        p.body = p.body.replace(new RegExp(`<(\\w+)[^>]*data-id="${c.target.slice(1)}"[^>]*>[\\s\\S]*?</\\1>`), c.content);
      } else return res.status(400).json({ error: { code: "20135", message: "bad patch" } });
    }
    res.status(204).end();
  });
  g.delete("/me/onenote/pages/:id", (req, res) => { if (!pages[req.params.id]) return notFound(res); delete pages[req.params.id]; res.status(204).end(); });
  g.post("/me/onenote/pages/:id/copyToSection", (req, res) => {
    const p = pages[req.params.id];
    if (!p) return notFound(res);
    const copy = { ...p, id: id("p"), section: req.body.id };
    pages[copy.id] = copy;
    const opId = id("op");
    ops[opId] = { polls: 0, resourceId: copy.id };
    res.status(202).json({ id: opId, status: "NotStarted" });
  });
  g.post("/me/onenote/sections/:id/copyToNotebook", (req, res) => {
    const s = sections[req.params.id];
    if (!s) return notFound(res);
    const copy = { id: id("s"), displayName: req.body.renameAs ?? s.displayName, parent: req.body.id };
    sections[copy.id] = copy;
    const opId = id("op");
    ops[opId] = { polls: 0, resourceId: copy.id };
    res.status(202).json({ id: opId, status: "NotStarted" });
  });
  g.get("/me/onenote/operations/:id", (req, res) => {
    const op = ops[req.params.id];
    if (!op) return notFound(res);
    op.polls++;
    res.json({ id: req.params.id, status: op.polls < 2 ? "Running" : "Completed", resourceId: op.resourceId });
  });
  g.get("/me/onenote/resources/:id/:value", (_req, res) => res.type("image/png").send(PNG));
  app.use("/v1.0", g);

  return new Promise<{ url: string; server: Server; state: typeof state; pages: typeof pages; sections: typeof sections }>((resolve) => {
    const server = app.listen(0, "127.0.0.1", () => {
      const port = (server.address() as any).port;
      resolve({ url: `http://127.0.0.1:${port}`, server, state, pages, sections });
    });
  });
}
