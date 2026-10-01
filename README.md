# OneNote connector for Claude

A hosted [MCP](https://modelcontextprotocol.io) server that lets Claude read, write, search and organise Microsoft OneNote.
Anyone can add it to Claude by pasting one URL and signing in with Microsoft. They don't install anything or register their own Azure app.

```
Claude ──OAuth (DCR + PKCE)──▶ this server ──OAuth──▶ Microsoft identity
   │                               │
   └──── MCP tool calls ─────────▶ └──── Microsoft Graph /me/onenote ───▶ OneNote
```

## What Claude can do

| Tool | What it does |
|---|---|
| `list_notebooks` | All notebooks, newest first |
| `get_notebook_structure` | Full tree of section groups (folders) and sections, with IDs |
| `list_pages` | Pages in a section, in OneNote order |
| `search_pages` | Title search across everything; optional full-text search within a notebook/section |
| `read_page` | Page as Markdown (to-dos, tables, image references), or as HTML with element IDs |
| `get_page_resource` | Fetch an embedded image so Claude can see it (diagrams, photos of the whiteboard) |
| `create_page` | New page from Markdown. `- [ ]` becomes real OneNote checkboxes |
| `append_to_page` | Add to the start or end of a page |
| `update_page` | Rename, or replace/insert specific paragraphs |
| `delete_page` | Delete a page (flagged destructive so Claude asks first) |
| `create_notebook` / `create_section` / `create_section_group` | Build structure |
| `move_page` | Move (or copy) a page to another section |
| `copy_section` | Copy a whole section into another notebook or folder, optionally renamed |
| `whoami` | Which Microsoft account is connected |

**What the Microsoft API can't do:** rename or delete notebooks and sections, or read handwriting/ink. Claude is told this, so it'll say so instead of failing quietly.

## Using it (end users)

1. In Claude, go to **Settings → Connectors → Add custom connector**.
2. Paste `https://<your-host>/mcp`.
3. Click **Connect**, review the consent screen, then sign in with Microsoft.

## Hosting it

### 1. Register an app in Microsoft Entra (about 5 minutes)

1. Go to <https://entra.microsoft.com> → **App registrations** → **New registration**.
2. **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*.
3. **Redirect URI:** platform **Web**, `https://<your-host>/oauth/microsoft/callback`.
4. In **Certificates & secrets**, create a client secret and copy the **Value**.
5. In **API permissions**, add Microsoft Graph → **Delegated**: `Notes.ReadWrite`, `User.Read`, `offline_access`.
6. Copy the **Application (client) ID** from Overview.

> **School and work accounts:** most organisations only let users consent to multi-tenant apps from a
> [verified publisher](https://learn.microsoft.com/entra/identity-platform/publisher-verification).
> Personal Microsoft accounts work right away. For students, either complete publisher verification
> (needs a Microsoft AI Cloud Partner Program ID) or ask the university's IT to approve the app for their tenant.

### 2. Configure and run

```bash
cp .env.example .env    # fill it in
openssl rand -hex 32    # → TOKEN_ENCRYPTION_KEY
npm ci && npm run build && npm start
```

or with Docker:

```bash
docker build -t onenote-mcp .
docker run -p 3000:3000 --env-file .env onenote-mcp
```

Set `DATABASE_URL` to a Postgres database for production. The server creates its one table (`kv`) on boot. Without a database it falls back to memory, and everyone has to reconnect after a restart.

The server is stateless, so you can run as many replicas as you like behind a load balancer. All state lives in Postgres.

## Security

- **Consent screen and CSRF:** every new client must be approved by the user before the server forwards them to Microsoft. This addresses the MCP spec's confused-deputy issue for proxy servers. The approval is bound to a `SameSite=Strict` cookie, so another site can't auto-submit it, and the page can't be framed.
- **Redirect URIs** must match the ones the client registered. PKCE (S256) is required.
- **Microsoft tokens never leave the server.** They're encrypted at rest with AES-256-GCM. Claude only gets opaque tokens issued by this server, stored as SHA-256 hashes, with refresh tokens rotated on every use.
- If the user revokes access at Microsoft, the stored sign-in is deleted. Claude gets a 401 on the next request and asks the user to reconnect.
- Only delegated `Notes.ReadWrite` is requested, so the server can only touch the signed-in user's own notebooks.

## Development

```bash
npm test                                         # end-to-end tests against a fake Microsoft + Graph
TEST_DATABASE_URL=postgres://... npm test         # same, against Postgres
npm run dev
```

The tests run the real OAuth flow end to end: discovery, dynamic registration, consent, Microsoft redirect, code exchange, refresh rotation and revocation. They then drive every tool through the official MCP client. They also cover Graph pagination, throttling retries and token expiry.

## Licence

MIT
