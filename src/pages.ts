export function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}

const STYLE = `
  :root { color-scheme: light dark; --bg:#f7f5f2; --card:#fff; --fg:#1d1b19; --muted:#6b655e; --accent:#7719aa; --border:#e6e1da; }
  @media (prefers-color-scheme: dark) { :root { --bg:#161514; --card:#201f1d; --fg:#ece8e3; --muted:#a39d95; --accent:#c487e8; --border:#33302c; } }
  * { box-sizing: border-box; }
  body { margin:0; min-height:100vh; display:grid; place-items:center; background:var(--bg); color:var(--fg);
         font: 16px/1.5 system-ui, -apple-system, "Segoe UI", sans-serif; padding:16px; }
  .card { width:100%; max-width:460px; background:var(--card); border:1px solid var(--border); border-radius:14px; padding:28px; }
  h1 { font-size:1.3rem; margin:0 0 8px; }
  p, li { color:var(--muted); }
  code { background:var(--bg); border:1px solid var(--border); padding:2px 6px; border-radius:6px; font-size:.9em; word-break:break-all; color:var(--fg); }
  .row { display:flex; gap:10px; margin-top:22px; }
  button { flex:1; font:inherit; padding:11px 14px; border-radius:9px; border:1px solid var(--border); background:transparent; color:var(--fg); cursor:pointer; }
  button.primary { background:var(--accent); border-color:var(--accent); color:#fff; font-weight:600; }
  .logo { width:40px; height:40px; border-radius:10px; background:var(--accent); color:#fff; display:grid; place-items:center; font-weight:700; margin-bottom:14px; }
`;

function shell(title: string, body: string): string {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title><style>${STYLE}</style></head><body><main class="card"><div class="logo">N</div>${body}</main></body></html>`;
}

export function consentPage(opts: { pendingId: string; csrf: string; clientName?: string; redirectUri: string }): string {
  const app = escapeHtml(opts.clientName || "An application");
  const host = escapeHtml(new URL(opts.redirectUri).host);
  return shell(
    "Connect OneNote",
    `<h1>${app} wants to access your OneNote</h1>
     <p>It will be able to read, create, edit, copy and delete notebooks, sections and pages in your OneNote.</p>
     <p>After you continue you'll sign in with Microsoft, and you'll be sent back to <code>${host}</code>.</p>
     <p>Only continue if you started this from Claude yourself.</p>
     <form method="post" action="/oauth/consent">
       <input type="hidden" name="pending_id" value="${escapeHtml(opts.pendingId)}">
       <input type="hidden" name="csrf" value="${escapeHtml(opts.csrf)}">
       <div class="row">
         <button type="submit" name="action" value="deny">Cancel</button>
         <button type="submit" name="action" value="allow" class="primary">Continue to Microsoft</button>
       </div>
     </form>`,
  );
}

export function messagePage(title: string, message: string): string {
  return shell(title, `<h1>${escapeHtml(title)}</h1><p>${escapeHtml(message)}</p>`);
}

export function landingPage(mcpUrl: string): string {
  return shell(
    "OneNote connector for Claude",
    `<h1>OneNote connector for Claude</h1>
     <p>Let Claude read, write, search and organise your OneNote notebooks.</p>
     <ol>
       <li>In Claude, open <b>Settings → Connectors</b> and choose <b>Add custom connector</b>.</li>
       <li>Paste this URL: <code>${escapeHtml(mcpUrl)}</code></li>
       <li>Click <b>Connect</b> and sign in with your Microsoft account.</li>
     </ol>
     <p>Works with personal Microsoft accounts and most work or school accounts. Some schools require IT approval first.</p>`,
  );
}
