import { loadConfig } from "./config.js";
import { MemoryStore, PostgresStore } from "./store.js";
import { createApp } from "./server.js";

const cfg = loadConfig();
const store = cfg.databaseUrl ? new PostgresStore(cfg.databaseUrl) : new MemoryStore();
if (!cfg.databaseUrl) console.warn("DATABASE_URL not set: using in-memory storage. Users will need to reconnect after every restart.");

const { app } = createApp(cfg, store);
app.listen(cfg.port, () => {
  console.log(`OneNote MCP server listening on :${cfg.port}`);
  console.log(`Connector URL: ${new URL("/mcp", cfg.baseUrl)}`);
  console.log(`Microsoft redirect URI to register: ${new URL("/oauth/microsoft/callback", cfg.baseUrl)}`);
});
