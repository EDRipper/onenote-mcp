import express from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { mcpAuthRouter, getOAuthProtectedResourceMetadataUrl } from "@modelcontextprotocol/sdk/server/auth/router.js";
import { requireBearerAuth } from "@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js";
import type { Config } from "./config.js";
import type { KVStore } from "./store.js";
import { OneNoteAuthProvider } from "./auth-provider.js";
import { Graph } from "./graph.js";
import { registerTools } from "./tools.js";
import { landingPage } from "./pages.js";

export const SERVER_INFO = { name: "onenote", version: "0.1.0" };

const INSTRUCTIONS = `Tools for the user's Microsoft OneNote.
Hierarchy: notebook → section groups (folders, can nest) → sections (tabs) → pages (and subpages).
Start with list_notebooks or get_notebook_structure to get IDs; IDs are required by every other tool.
Before bulk reorganising (moving many pages, copying sections), show the user the plan and get confirmation.
The API cannot rename or delete notebooks/sections, or read handwriting/ink; say so rather than guessing.
For diagrams use insert_diagram, or fenced \`\`\`mermaid / \`\`\`svg blocks inside Markdown page content; they become images on the page.`;

export function createApp(cfg: Config, store: KVStore) {
  const provider = new OneNoteAuthProvider(cfg, store);
  const mcpUrl = new URL("/mcp", cfg.baseUrl);
  const app = express();
  app.set("trust proxy", 1);
  app.disable("x-powered-by");

  app.get("/", (_req, res) => res.type("html").send(landingPage(mcpUrl.toString())));
  app.get("/health", (_req, res) => res.json({ ok: true }));

  app.use(
    mcpAuthRouter({
      provider,
      issuerUrl: cfg.baseUrl,
      resourceServerUrl: mcpUrl,
      resourceName: "OneNote",
      serviceDocumentationUrl: cfg.baseUrl,
    }),
  );
  app.post("/oauth/consent", express.urlencoded({ extended: false }), provider.handleConsent);
  app.get("/oauth/microsoft/callback", provider.handleMicrosoftCallback);

  const graphFor = (authInfo: { extra?: Record<string, unknown> } | undefined) => {
    const grantId = authInfo?.extra?.grantId;
    if (typeof grantId !== "string") throw new Error("Not authenticated");
    return new Graph(cfg.graphBase, (force) => provider.getMicrosoftToken(grantId, force));
  };

  const auth = requireBearerAuth({ verifier: provider, resourceMetadataUrl: getOAuthProtectedResourceMetadataUrl(mcpUrl) });

  // Stateless Streamable HTTP: a fresh server per request scales horizontally with no sticky sessions.
  app.post("/mcp", auth, express.json({ limit: "8mb" }), async (req, res) => {
    const server = new McpServer(SERVER_INFO, { instructions: INSTRUCTIONS });
    registerTools(server, graphFor);
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => {
      transport.close();
      server.close();
    });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err) {
      console.error("MCP request failed:", err);
      if (!res.headersSent) res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal error" }, id: null });
    }
  });
  const methodNotAllowed: express.RequestHandler = (_req, res) => {
    res.status(405).set("Allow", "POST").json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed" }, id: null });
  };
  app.get("/mcp", auth, methodNotAllowed);
  app.delete("/mcp", auth, methodNotAllowed);

  return { app, provider };
}
