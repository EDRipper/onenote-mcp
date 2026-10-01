import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { Graph } from "./graph.js";
import { buildPageHtml, htmlToMarkdown, prepareContent, prepareMarkdown, type PageImage } from "./convert.js";

const ON = "/me/onenote";
const MAX_PAGE_CHARS = 100_000;
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;

type GraphFactory = (authInfo: { extra?: Record<string, unknown> } | undefined) => Graph;

const text = (s: string): CallToolResult => ({ content: [{ type: "text", text: s }] });
const json = (v: unknown): CallToolResult => text(JSON.stringify(v, null, 2));

/** Wrap a handler so thrown errors become readable tool errors instead of protocol errors. */
function safe<A>(fn: (args: A, extra: any) => Promise<CallToolResult>) {
  return async (args: A, extra: any): Promise<CallToolResult> => {
    try {
      return await fn(args, extra);
    } catch (err) {
      return { isError: true, content: [{ type: "text", text: err instanceof Error ? err.message : String(err) }] };
    }
  };
}

const enc = encodeURIComponent;
const odataString = (s: string) => `'${s.replace(/'/g, "''")}'`;
const formatSchema = z.enum(["markdown", "html"]).default("markdown")
  .describe("Format of `content`. Markdown supports headings, lists, tables, code, links and `- [ ]` to-do checkboxes.");

const DIAGRAM_HELP =
  " Diagrams: put a fenced ```mermaid block (flowcharts, sequence/state/class diagrams, Gantt, etc.) or a fenced ```svg block " +
  "(full <svg> markup: circuits, free-body diagrams, labelled sketches, plots) in Markdown content and it is rendered to an image on the page. " +
  "Text after the language becomes alt text, e.g. ```mermaid Control loop.";

/** POST a new page, as multipart when it carries rendered images. */
async function postPage(g: Graph, sectionId: string, title: string, html: string, images: PageImage[]) {
  const page = buildPageHtml(title, html);
  const path = `${ON}/sections/${encodeURIComponent(sectionId)}/pages`;
  if (!images.length) return g.request("POST", path, { body: page, contentType: "application/xhtml+xml" });
  const form = new FormData();
  form.append("Presentation", new Blob([page], { type: "text/html" }));
  for (const img of images) form.append(img.name, new Blob([new Uint8Array(img.png)], { type: "image/png" }), `${img.name}.png`);
  return g.request("POST", path, { body: form });
}

/** PATCH page content, as multipart when commands reference rendered images. */
async function patchPage(g: Graph, pageId: string, commands: unknown[], images: PageImage[]) {
  const path = `${ON}/pages/${encodeURIComponent(pageId)}/content`;
  if (!images.length) return g.request("PATCH", path, { body: commands });
  const form = new FormData();
  form.append("Commands", new Blob([JSON.stringify(commands)], { type: "application/json" }));
  for (const img of images) form.append(img.name, new Blob([new Uint8Array(img.png)], { type: "image/png" }), `${img.name}.png`);
  return g.request("PATCH", path, { body: form });
}

interface Container { kind: "notebook" | "sectionGroup"; id: string }
function container(notebookId?: string, sectionGroupId?: string): Container {
  if (!!notebookId === !!sectionGroupId) throw new Error("Provide exactly one of notebook_id or section_group_id.");
  return notebookId ? { kind: "notebook", id: notebookId } : { kind: "sectionGroup", id: sectionGroupId! };
}
const containerPath = (c: Container) => `${ON}/${c.kind === "notebook" ? "notebooks" : "sectionGroups"}/${enc(c.id)}`;

// ---------- structure helpers ----------

interface TreeNode { type: "section" | "sectionGroup"; id: string; name: string; children?: TreeNode[] }

async function loadTree(g: Graph, c: Container, depth = 0): Promise<TreeNode[]> {
  const base = containerPath(c);
  const [sections, groups] = await Promise.all([
    g.list<any>(`${base}/sections`, { $select: "id,displayName", $orderby: "displayName" }),
    g.list<any>(`${base}/sectionGroups`, { $select: "id,displayName", $orderby: "displayName" }),
  ]);
  const groupNodes: TreeNode[] = await Promise.all(
    groups.map(async (sg) => ({
      type: "sectionGroup" as const,
      id: sg.id,
      name: sg.displayName,
      children: depth < 6 ? await loadTree(g, { kind: "sectionGroup", id: sg.id }, depth + 1) : [],
    })),
  );
  return [...groupNodes, ...sections.map((s) => ({ type: "section" as const, id: s.id, name: s.displayName }))];
}

function renderTree(nodes: TreeNode[], indent = "  "): string {
  return nodes
    .map((n) =>
      n.type === "section"
        ? `${indent}📄 ${n.name}  [section_id: ${n.id}]`
        : `${indent}📁 ${n.name}  [section_group_id: ${n.id}]\n${renderTree(n.children ?? [], indent + "  ")}`,
    )
    .join("\n");
}

function flattenSections(nodes: TreeNode[], path: string[] = []): { id: string; path: string }[] {
  return nodes.flatMap((n) =>
    n.type === "section" ? [{ id: n.id, path: [...path, n.name].join(" / ") }] : flattenSections(n.children ?? [], [...path, n.name]),
  );
}

const PAGE_SELECT = "id,title,createdDateTime,lastModifiedDateTime,level,order";
const pageSummary = (p: any) => ({
  page_id: p.id,
  title: p.title || "(untitled)",
  last_modified: p.lastModifiedDateTime,
  ...(p.parentNotebook ? { notebook: p.parentNotebook.displayName } : {}),
  ...(p.parentSection ? { section: p.parentSection.displayName, section_id: p.parentSection.id } : {}),
  ...(p.level ? { subpage_level: p.level } : {}),
});

async function waitForOperation(g: Graph, op: any, timeoutMs = 45_000): Promise<any> {
  const id = op?.id;
  if (!id) throw new Error("Microsoft did not return an operation ID for the copy.");
  const start = Date.now();
  let delay = 500;
  while (Date.now() - start < timeoutMs) {
    const status = await g.get(`${ON}/operations/${enc(id)}`);
    if (status.status === "Completed") return status;
    if (status.status === "Failed") throw new Error(`Copy failed: ${JSON.stringify(status.error ?? status)}`);
    await new Promise((r) => setTimeout(r, delay));
    delay = Math.min(delay * 1.5, 3000);
  }
  throw new Error(`Copy is still running after ${timeoutMs / 1000}s (operation ${id}). It will probably finish; check the target section shortly.`);
}

// ---------- tools ----------

export function registerTools(server: McpServer, graphFor: GraphFactory) {
  const RO = { readOnlyHint: true, openWorldHint: true } as const;
  const WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;

  server.registerTool(
    "whoami",
    { title: "Who am I signed in as", description: "Show which Microsoft account this OneNote connection is using.", inputSchema: {}, annotations: RO },
    safe(async (_a, extra) => {
      const me = await graphFor(extra.authInfo).get("/me", { $select: "displayName,mail,userPrincipalName" });
      return json({ name: me.displayName, email: me.mail ?? me.userPrincipalName });
    }),
  );

  server.registerTool(
    "list_notebooks",
    {
      title: "List notebooks",
      description: "List all OneNote notebooks the user can access, with IDs. Start here.",
      inputSchema: {},
      annotations: RO,
    },
    safe(async (_a, extra) => {
      const nbs = await graphFor(extra.authInfo).list<any>(`${ON}/notebooks`, {
        $select: "id,displayName,lastModifiedDateTime,isShared,userRole",
        $orderby: "lastModifiedDateTime desc",
      });
      return json(nbs.map((n) => ({ notebook_id: n.id, name: n.displayName, last_modified: n.lastModifiedDateTime, shared: n.isShared, role: n.userRole })));
    }),
  );

  server.registerTool(
    "get_notebook_structure",
    {
      title: "Show notebook structure",
      description:
        "Show the full tree of section groups (folders) and sections inside a notebook, with IDs. " +
        "Omit notebook_id to show every notebook. Use this before reorganising.",
      inputSchema: { notebook_id: z.string().optional().describe("Notebook ID from list_notebooks. Omit for all notebooks.") },
      annotations: RO,
    },
    safe(async ({ notebook_id }: { notebook_id?: string }, extra) => {
      const g = graphFor(extra.authInfo);
      const nbs = notebook_id
        ? [await g.get(`${ON}/notebooks/${enc(notebook_id)}`, { $select: "id,displayName" })]
        : await g.list<any>(`${ON}/notebooks`, { $select: "id,displayName", $orderby: "displayName" });
      const parts = await Promise.all(
        nbs.map(async (nb) => {
          const tree = await loadTree(g, { kind: "notebook", id: nb.id });
          return `📓 ${nb.displayName}  [notebook_id: ${nb.id}]\n${renderTree(tree) || "  (empty)"}`;
        }),
      );
      return text(parts.join("\n\n"));
    }),
  );

  server.registerTool(
    "list_pages",
    {
      title: "List pages in a section",
      description: "List the pages in a section (in their OneNote order), with IDs and last-modified times.",
      inputSchema: {
        section_id: z.string().describe("Section ID from get_notebook_structure."),
        limit: z.number().int().min(1).max(500).default(100),
      },
      annotations: RO,
    },
    safe(async ({ section_id, limit }: { section_id: string; limit: number }, extra) => {
      const pages = await graphFor(extra.authInfo).list<any>(
        `${ON}/sections/${enc(section_id)}/pages`,
        { $select: PAGE_SELECT, $top: Math.min(limit, 100), pagelevel: "true" },
        limit,
      );
      pages.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
      return json(pages.map(pageSummary));
    }),
  );

  server.registerTool(
    "search_pages",
    {
      title: "Search pages",
      description:
        "Find pages by title across all notebooks. Set search_content=true together with a section_id or notebook_id " +
        "to also search inside page text (slower; scans up to `max_scan` pages).",
      inputSchema: {
        query: z.string().min(1),
        notebook_id: z.string().optional().describe("Limit to one notebook."),
        section_id: z.string().optional().describe("Limit to one section."),
        search_content: z.boolean().default(false),
        max_scan: z.number().int().min(1).max(200).default(60),
      },
      annotations: RO,
    },
    safe(async (a: { query: string; notebook_id?: string; section_id?: string; search_content: boolean; max_scan: number }, extra) => {
      const g = graphFor(extra.authInfo);
      const q = a.query.toLowerCase();
      const expand = "parentNotebook($select=displayName),parentSection($select=id,displayName)";

      // Sections in scope
      let sections: { id: string; path: string }[] | undefined;
      if (a.section_id) sections = [{ id: a.section_id, path: "" }];
      else if (a.notebook_id) sections = flattenSections(await loadTree(g, { kind: "notebook", id: a.notebook_id }));

      // 1) Title matches
      let pages: any[];
      if (sections) {
        const lists = await Promise.all(
          sections.map((s) => g.list<any>(`${ON}/sections/${enc(s.id)}/pages`, { $select: PAGE_SELECT, $expand: expand, $top: 100 }, 500)),
        );
        pages = lists.flat();
      } else {
        pages = await g.list<any>(`${ON}/pages`, {
          $select: PAGE_SELECT,
          $expand: expand,
          $filter: `contains(tolower(title),${odataString(q)})`,
          $top: 100,
        }, 200);
      }
      const titleHits = pages.filter((p) => (p.title ?? "").toLowerCase().includes(q));

      // 2) Content matches
      const contentHits: any[] = [];
      if (a.search_content) {
        if (!sections) throw new Error("search_content needs a section_id or notebook_id to keep the scan bounded.");
        const toScan = pages
          .filter((p) => !titleHits.includes(p))
          .sort((x, y) => String(y.lastModifiedDateTime).localeCompare(String(x.lastModifiedDateTime)))
          .slice(0, a.max_scan);
        for (let i = 0; i < toScan.length; i += 6) {
          await Promise.all(
            toScan.slice(i, i + 6).map(async (p) => {
              const res: Response = await g.request("GET", `${ON}/pages/${enc(p.id)}/content`, { raw: true });
              const md = htmlToMarkdown(await res.text());
              const idx = md.toLowerCase().indexOf(q);
              if (idx >= 0) {
                contentHits.push({ ...pageSummary(p), snippet: "…" + md.slice(Math.max(0, idx - 120), idx + q.length + 120).replace(/\s+/g, " ") + "…" });
              }
            }),
          );
        }
      }
      return json({
        title_matches: titleHits.map(pageSummary),
        ...(a.search_content ? { content_matches: contentHits, pages_scanned: Math.min(pages.length - titleHits.length, a.max_scan) } : {}),
      });
    }),
  );

  server.registerTool(
    "read_page",
    {
      title: "Read a page",
      description:
        "Read a page's content. Markdown (default) is best for reading; images appear as `onenote-resource:<id>` links " +
        "which get_page_resource can fetch. Use format=html to get element IDs (data-id) for targeted edits with update_page. " +
        "Handwriting/ink is not included by the OneNote API.",
      inputSchema: {
        page_id: z.string(),
        format: z.enum(["markdown", "html"]).default("markdown"),
      },
      annotations: RO,
    },
    safe(async ({ page_id, format }: { page_id: string; format: "markdown" | "html" }, extra) => {
      const g = graphFor(extra.authInfo);
      const [meta, res] = await Promise.all([
        g.get(`${ON}/pages/${enc(page_id)}`, { $select: "id,title,lastModifiedDateTime", $expand: "parentNotebook($select=displayName),parentSection($select=id,displayName)" }),
        g.request<Response>("GET", `${ON}/pages/${enc(page_id)}/content`, { raw: true, query: format === "html" ? { includeIDs: "true" } : undefined }),
      ]);
      const html = await res.text();
      let body = format === "html" ? html : htmlToMarkdown(html);
      if (body.length > MAX_PAGE_CHARS) body = body.slice(0, MAX_PAGE_CHARS) + `\n\n[truncated — page is ${body.length} characters]`;
      const header = `# ${meta.title || "(untitled)"}\n_${meta.parentNotebook?.displayName ?? ""} / ${meta.parentSection?.displayName ?? ""} · last modified ${meta.lastModifiedDateTime}_\n\n`;
      return text(header + body);
    }),
  );

  server.registerTool(
    "get_page_resource",
    {
      title: "Get an image or attachment",
      description: "Fetch an image or file embedded in a page, using the ID from an `onenote-resource:<id>` link in read_page. Images are returned so you can see them.",
      inputSchema: { resource_id: z.string() },
      annotations: RO,
    },
    safe(async ({ resource_id }: { resource_id: string }, extra) => {
      const res = await graphFor(extra.authInfo).request<Response>("GET", `${ON}/resources/${enc(resource_id)}/$value`, { raw: true });
      const type = (res.headers.get("Content-Type") ?? "application/octet-stream").split(";")[0];
      const buf = Buffer.from(await res.arrayBuffer());
      if (type.startsWith("image/") && buf.length <= MAX_IMAGE_BYTES) {
        return { content: [{ type: "image", data: buf.toString("base64"), mimeType: type }] };
      }
      if (type.startsWith("text/") && buf.length <= MAX_PAGE_CHARS) return text(buf.toString("utf8"));
      return text(`Resource is ${type}, ${buf.length} bytes — too large or not viewable here.`);
    }),
  );

  server.registerTool(
    "create_page",
    {
      title: "Create a page",
      description: "Create a new page in a section." + DIAGRAM_HELP,
      inputSchema: {
        section_id: z.string(),
        title: z.string(),
        content: z.string().default("").describe("Page body."),
        format: formatSchema,
      },
      annotations: WRITE,
    },
    safe(async (a: { section_id: string; title: string; content: string; format: "markdown" | "html" }, extra) => {
      const { html, images } = await prepareContent(a.content, a.format);
      const page = await postPage(graphFor(extra.authInfo), a.section_id, a.title, html, images);
      return json({ created: true, page_id: page.id, title: page.title, diagrams: images.length || undefined, web_url: page.links?.oneNoteWebUrl?.href });
    }),
  );

  server.registerTool(
    "append_to_page",
    {
      title: "Add to a page",
      description: "Add content to the end (or start) of an existing page without touching what's already there." + DIAGRAM_HELP,
      inputSchema: {
        page_id: z.string(),
        content: z.string(),
        format: formatSchema,
        position: z.enum(["end", "start"]).default("end"),
      },
      annotations: WRITE,
    },
    safe(async (a: { page_id: string; content: string; format: "markdown" | "html"; position: "end" | "start" }, extra) => {
      const { html, images } = await prepareContent(a.content, a.format);
      await patchPage(graphFor(extra.authInfo), a.page_id, [{ target: "body", action: a.position === "end" ? "append" : "prepend", content: html }], images);
      return text(images.length ? `Added to page (${images.length} diagram${images.length > 1 ? "s" : ""}).` : "Added to page.");
    }),
  );

  server.registerTool(
    "update_page",
    {
      title: "Edit or rename a page",
      description:
        "Rename a page and/or make targeted edits. For edits, call read_page with format=html first and target elements by their " +
        "data-id as `#<data-id>` (or by `id` attribute). Actions: replace, append (inside the element, at end), prepend, insert (before/after)." + DIAGRAM_HELP,
      inputSchema: {
        page_id: z.string(),
        title: z.string().optional().describe("New page title."),
        edits: z
          .array(
            z.object({
              target: z.string().describe("`#data-id`, element id, or `body`."),
              action: z.enum(["replace", "append", "prepend", "insert"]),
              position: z.enum(["before", "after"]).optional().describe("Only for insert/append."),
              content: z.string(),
              format: z.enum(["markdown", "html"]).default("markdown"),
            }),
          )
          .optional(),
      },
      annotations: { ...WRITE, destructiveHint: true },
    },
    safe(async (a: { page_id: string; title?: string; edits?: any[] }, extra) => {
      const commands: any[] = [];
      const images: PageImage[] = [];
      if (a.title !== undefined) commands.push({ target: "title", action: "replace", content: a.title });
      for (const [i, e] of (a.edits ?? []).entries()) {
        const prepared = await prepareContent(e.content, e.format, `edit${i + 1}diagram`);
        images.push(...prepared.images);
        commands.push({ target: e.target, action: e.action, ...(e.position ? { position: e.position } : {}), content: prepared.html });
      }
      if (!commands.length) throw new Error("Nothing to change: give a title and/or edits.");
      await patchPage(graphFor(extra.authInfo), a.page_id, commands, images);
      return text(`Applied ${commands.length} change(s).`);
    }),
  );

  server.registerTool(
    "insert_diagram",
    {
      title: "Draw a diagram",
      description:
        "Render a diagram and put it in OneNote, either at the end/start of an existing page (page_id) or as a new page (section_id + title). " +
        "kind=mermaid for flowcharts, sequence, state, class, ER, Gantt, mindmaps, etc. " +
        "kind=svg for anything Mermaid can't draw (circuits, free-body diagrams, mechanisms, annotated graphs): pass a complete <svg> with width/height or a viewBox; " +
        "use plain shapes, paths and <text> (no scripts, external images or web fonts). Rendering errors are returned so you can fix the source and retry.",
      inputSchema: {
        kind: z.enum(["mermaid", "svg"]),
        source: z.string().describe("Mermaid code, or full SVG markup."),
        caption: z.string().optional().describe("Shown in italics under the diagram; also used as alt text."),
        page_id: z.string().optional().describe("Add to this existing page."),
        position: z.enum(["end", "start"]).default("end"),
        section_id: z.string().optional().describe("Or create a new page in this section."),
        title: z.string().optional().describe("Title for the new page (with section_id)."),
      },
      annotations: WRITE,
    },
    safe(async (a: { kind: "mermaid" | "svg"; source: string; caption?: string; page_id?: string; position: "end" | "start"; section_id?: string; title?: string }, extra) => {
      if (!!a.page_id === !!a.section_id) throw new Error("Provide exactly one of page_id (add to a page) or section_id (new page).");
      const fence = "~~~~~~~~";
      const markdown = `${fence}${a.kind} ${a.caption ?? ""}\n${a.source}\n${fence}${a.caption ? `\n\n*${a.caption}*` : ""}`;
      const { html, images } = await prepareMarkdown(markdown);
      const g = graphFor(extra.authInfo);
      if (a.page_id) {
        await patchPage(g, a.page_id, [{ target: "body", action: a.position === "end" ? "append" : "prepend", content: html }], images);
        return text("Diagram added to page.");
      }
      const page = await postPage(g, a.section_id!, a.title ?? a.caption ?? "Diagram", html, images);
      return json({ created: true, page_id: page.id, title: page.title, web_url: page.links?.oneNoteWebUrl?.href });
    }),
  );

  server.registerTool(
    "delete_page",
    {
      title: "Delete a page",
      description: "Permanently delete a page. It goes to OneNote's 'Deleted Pages' for a while but can't be restored through this connector.",
      inputSchema: { page_id: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
    },
    safe(async ({ page_id }: { page_id: string }, extra) => {
      await graphFor(extra.authInfo).request("DELETE", `${ON}/pages/${enc(page_id)}`);
      return text("Page deleted.");
    }),
  );

  server.registerTool(
    "create_notebook",
    {
      title: "Create a notebook",
      description: "Create a new notebook in the user's OneDrive.",
      inputSchema: { name: z.string().max(128) },
      annotations: WRITE,
    },
    safe(async ({ name }: { name: string }, extra) => {
      const nb = await graphFor(extra.authInfo).request("POST", `${ON}/notebooks`, { body: { displayName: name } });
      return json({ created: true, notebook_id: nb.id, name: nb.displayName });
    }),
  );

  server.registerTool(
    "create_section",
    {
      title: "Create a section",
      description: "Create a section (tab) inside a notebook or section group. Provide exactly one of notebook_id or section_group_id.",
      inputSchema: { name: z.string().max(50), notebook_id: z.string().optional(), section_group_id: z.string().optional() },
      annotations: WRITE,
    },
    safe(async (a: { name: string; notebook_id?: string; section_group_id?: string }, extra) => {
      const c = container(a.notebook_id, a.section_group_id);
      const s = await graphFor(extra.authInfo).request("POST", `${containerPath(c)}/sections`, { body: { displayName: a.name } });
      return json({ created: true, section_id: s.id, name: s.displayName });
    }),
  );

  server.registerTool(
    "create_section_group",
    {
      title: "Create a section group (folder)",
      description: "Create a section group — a folder for sections — inside a notebook or another section group. Provide exactly one of notebook_id or section_group_id.",
      inputSchema: { name: z.string().max(50), notebook_id: z.string().optional(), section_group_id: z.string().optional() },
      annotations: WRITE,
    },
    safe(async (a: { name: string; notebook_id?: string; section_group_id?: string }, extra) => {
      const c = container(a.notebook_id, a.section_group_id);
      const sg = await graphFor(extra.authInfo).request("POST", `${containerPath(c)}/sectionGroups`, { body: { displayName: a.name } });
      return json({ created: true, section_group_id: sg.id, name: sg.displayName });
    }),
  );

  server.registerTool(
    "move_page",
    {
      title: "Move or copy a page",
      description:
        "Move a page to another section (copies it, waits for the copy to finish, then deletes the original). " +
        "Set keep_original=true to copy instead. The moved page gets a new page_id.",
      inputSchema: { page_id: z.string(), target_section_id: z.string(), keep_original: z.boolean().default(false) },
      annotations: { ...WRITE, destructiveHint: true },
    },
    safe(async (a: { page_id: string; target_section_id: string; keep_original: boolean }, extra) => {
      const g = graphFor(extra.authInfo);
      const op = await g.request("POST", `${ON}/pages/${enc(a.page_id)}/copyToSection`, { body: { id: a.target_section_id } });
      const done = await waitForOperation(g, op);
      if (!a.keep_original) await g.request("DELETE", `${ON}/pages/${enc(a.page_id)}`);
      return json({ [a.keep_original ? "copied" : "moved"]: true, new_page: done.resourceId ?? done.resourceLocation });
    }),
  );

  server.registerTool(
    "copy_section",
    {
      title: "Copy a section",
      description:
        "Copy a whole section (with its pages) into a notebook or section group, optionally renaming it. " +
        "The Microsoft API can't delete or rename sections in place, so to 'move' a section copy it here and then ask the user to delete the original in OneNote.",
      inputSchema: {
        section_id: z.string(),
        target_notebook_id: z.string().optional(),
        target_section_group_id: z.string().optional(),
        new_name: z.string().max(50).optional(),
      },
      annotations: WRITE,
    },
    safe(async (a: { section_id: string; target_notebook_id?: string; target_section_group_id?: string; new_name?: string }, extra) => {
      const c = container(a.target_notebook_id, a.target_section_group_id);
      const g = graphFor(extra.authInfo);
      const action = c.kind === "notebook" ? "copyToNotebook" : "copyToSectionGroup";
      const op = await g.request("POST", `${ON}/sections/${enc(a.section_id)}/${action}`, {
        body: { id: c.id, ...(a.new_name ? { renameAs: a.new_name } : {}) },
      });
      const done = await waitForOperation(g, op, 90_000);
      return json({ copied: true, new_section: done.resourceId ?? done.resourceLocation });
    }),
  );
}
