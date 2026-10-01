function required(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required environment variable ${name}`);
  return v;
}

export interface Config {
  port: number;
  /** Public base URL of this server, e.g. https://onenote-mcp.example.com */
  baseUrl: URL;
  msClientId: string;
  msClientSecret: string;
  /** "common" = personal Microsoft accounts + work/school accounts */
  msTenant: string;
  msAuthority: string;
  graphBase: string;
  msScopes: string[];
  /** 32-byte key (hex or base64) used to encrypt Microsoft tokens at rest */
  encryptionKey: Buffer;
  databaseUrl?: string;
}

function parseKey(raw: string): Buffer {
  const hex = /^[0-9a-fA-F]{64}$/.test(raw) ? Buffer.from(raw, "hex") : null;
  const key = hex ?? Buffer.from(raw, "base64");
  if (key.length !== 32) {
    throw new Error("TOKEN_ENCRYPTION_KEY must be 32 bytes (64 hex chars or base64). Generate one with: openssl rand -hex 32");
  }
  return key;
}

export function loadConfig(): Config {
  const extra = (process.env.MS_EXTRA_SCOPES ?? "").split(/[\s,]+/).filter(Boolean);
  return {
    port: Number(process.env.PORT ?? 3000),
    baseUrl: new URL(required("BASE_URL")),
    msClientId: required("MS_CLIENT_ID"),
    msClientSecret: required("MS_CLIENT_SECRET"),
    msTenant: process.env.MS_TENANT ?? "common",
    msAuthority: (process.env.MS_AUTHORITY ?? "https://login.microsoftonline.com").replace(/\/$/, ""),
    graphBase: (process.env.GRAPH_BASE ?? "https://graph.microsoft.com/v1.0").replace(/\/$/, ""),
    msScopes: ["offline_access", "User.Read", "Notes.ReadWrite", ...extra],
    encryptionKey: parseKey(required("TOKEN_ENCRYPTION_KEY")),
    databaseUrl: process.env.DATABASE_URL || undefined,
  };
}
