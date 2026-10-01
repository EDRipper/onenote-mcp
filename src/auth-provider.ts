import type { Request, Response } from "express";
import type { OAuthServerProvider, AuthorizationParams } from "@modelcontextprotocol/sdk/server/auth/provider.js";
import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js";
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js";
import type { OAuthClientInformationFull, OAuthTokens } from "@modelcontextprotocol/sdk/shared/auth.js";
import { InvalidGrantError, InvalidTokenError } from "@modelcontextprotocol/sdk/server/auth/errors.js";
import type { Config } from "./config.js";
import type { KVStore } from "./store.js";
import { MicrosoftOAuth, MicrosoftOAuthError } from "./microsoft.js";
import { decrypt, encrypt, randomToken, sha256, sha256Challenge } from "./crypto.js";
import { consentPage, messagePage } from "./pages.js";

const ACCESS_TOKEN_TTL = 60 * 60; // 1 hour
const REFRESH_TOKEN_TTL = 90 * 24 * 60 * 60; // 90 days, rolling
const GRANT_TTL = REFRESH_TOKEN_TTL;
const PENDING_TTL = 10 * 60;
const CODE_TTL = 5 * 60;

interface PendingAuthorization {
  clientId: string;
  clientName?: string;
  redirectUri: string;
  codeChallenge: string;
  state?: string;
  scopes: string[];
  resource?: string;
  csrf: string;
}

interface MicrosoftLogin {
  pending: PendingAuthorization;
  msVerifier: string;
}

interface AuthCode {
  clientId: string;
  codeChallenge: string;
  redirectUri: string;
  grantId: string;
  scopes: string[];
  resource?: string;
}

interface TokenRecord {
  grantId: string;
  clientId: string;
  scopes: string[];
  resource?: string;
  expiresAt?: number;
}

/** Stored encrypted. One grant per user sign-in. */
export interface Grant {
  msRefreshToken: string;
  msAccessToken: string;
  msExpiresAt: number; // epoch ms
}

export class ReauthRequiredError extends Error {}

export class OneNoteAuthProvider implements OAuthServerProvider {
  readonly ms: MicrosoftOAuth;
  private refreshLocks = new Map<string, Promise<string>>();

  constructor(private cfg: Config, private store: KVStore) {
    this.ms = new MicrosoftOAuth(cfg);
  }

  get clientsStore(): OAuthRegisteredClientsStore {
    const store = this.store;
    return {
      getClient: (clientId) => store.get<OAuthClientInformationFull>(`client:${clientId}`),
      registerClient: async (client) => {
        const full = client as OAuthClientInformationFull;
        await store.set(`client:${full.client_id}`, full);
        return full;
      },
    };
  }

  // ---- Step 1: Claude sends the user here. Show our consent screen. ----
  async authorize(client: OAuthClientInformationFull, params: AuthorizationParams, res: Response): Promise<void> {
    const id = randomToken(24);
    const pending: PendingAuthorization = {
      clientId: client.client_id,
      clientName: client.client_name,
      redirectUri: params.redirectUri,
      codeChallenge: params.codeChallenge,
      state: params.state,
      scopes: params.scopes ?? [],
      resource: params.resource?.toString(),
      csrf: randomToken(24),
    };
    await this.store.set(`pending:${id}`, pending, PENDING_TTL);
    res.setHeader("Set-Cookie", `onenote_mcp_csrf=${pending.csrf}; Path=/oauth; HttpOnly; Secure; SameSite=Strict; Max-Age=${PENDING_TTL}`);
    res.setHeader("X-Frame-Options", "DENY");
    res.setHeader("Content-Security-Policy", "frame-ancestors 'none'");
    res.status(200).type("html").send(
      consentPage({ pendingId: id, csrf: pending.csrf, clientName: client.client_name, redirectUri: params.redirectUri }),
    );
  }

  // ---- Step 2: user clicked Allow / Cancel. ----
  handleConsent = async (req: Request, res: Response) => {
    const { pending_id, csrf, action } = req.body ?? {};
    const cookie = parseCookies(req.headers.cookie)["onenote_mcp_csrf"];
    const pending = typeof pending_id === "string" ? await this.store.take<PendingAuthorization>(`pending:${pending_id}`) : undefined;
    if (!pending || !csrf || csrf !== pending.csrf || cookie !== pending.csrf) {
      res.status(400).type("html").send(messagePage("Link expired", "This sign-in link has expired. Go back to Claude and try connecting again."));
      return;
    }
    if (action !== "allow") {
      res.redirect(302, withParams(pending.redirectUri, { error: "access_denied", state: pending.state }));
      return;
    }
    const msState = randomToken(24);
    const msVerifier = randomToken(48);
    await this.store.set(`mslogin:${msState}`, { pending, msVerifier } satisfies MicrosoftLogin, PENDING_TTL);
    res.redirect(302, this.ms.authorizeUrl(msState, sha256Challenge(msVerifier)));
  };

  // ---- Step 3: Microsoft redirects back here. ----
  handleMicrosoftCallback = async (req: Request, res: Response) => {
    const state = String(req.query.state ?? "");
    const login = state ? await this.store.take<MicrosoftLogin>(`mslogin:${state}`) : undefined;
    if (!login) {
      res.status(400).type("html").send(messagePage("Link expired", "This sign-in link has expired. Go back to Claude and try connecting again."));
      return;
    }
    const { pending } = login;
    if (req.query.error) {
      res.redirect(302, withParams(pending.redirectUri, {
        error: "access_denied",
        error_description: String(req.query.error_description ?? req.query.error),
        state: pending.state,
      }));
      return;
    }
    try {
      const tokens = await this.ms.exchangeCode(String(req.query.code ?? ""), login.msVerifier);
      if (!tokens.refresh_token) throw new MicrosoftOAuthError("no_refresh_token", "Microsoft did not return a refresh token (is offline_access granted?)");
      const grantId = randomToken(24);
      await this.saveGrant(grantId, {
        msRefreshToken: tokens.refresh_token,
        msAccessToken: tokens.access_token,
        msExpiresAt: Date.now() + tokens.expires_in * 1000,
      });
      const code = randomToken(32);
      const authCode: AuthCode = {
        clientId: pending.clientId,
        codeChallenge: pending.codeChallenge,
        redirectUri: pending.redirectUri,
        grantId,
        scopes: pending.scopes,
        resource: pending.resource,
      };
      await this.store.set(`code:${sha256(code)}`, authCode, CODE_TTL);
      res.redirect(302, withParams(pending.redirectUri, { code, state: pending.state }));
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error("Microsoft token exchange failed:", msg);
      res.redirect(302, withParams(pending.redirectUri, { error: "server_error", error_description: msg, state: pending.state }));
    }
  };

  // ---- Step 4: Claude exchanges our code for our tokens. ----
  async challengeForAuthorizationCode(client: OAuthClientInformationFull, authorizationCode: string): Promise<string> {
    const code = await this.store.get<AuthCode>(`code:${sha256(authorizationCode)}`);
    if (!code || code.clientId !== client.client_id) throw new InvalidGrantError("Invalid authorization code");
    return code.codeChallenge;
  }

  async exchangeAuthorizationCode(
    client: OAuthClientInformationFull,
    authorizationCode: string,
    _codeVerifier?: string,
    redirectUri?: string,
  ): Promise<OAuthTokens> {
    const code = await this.store.take<AuthCode>(`code:${sha256(authorizationCode)}`);
    if (!code || code.clientId !== client.client_id) throw new InvalidGrantError("Invalid authorization code");
    if (redirectUri && redirectUri !== code.redirectUri) throw new InvalidGrantError("redirect_uri mismatch");
    return this.issueTokens(code.grantId, client.client_id, code.scopes, code.resource);
  }

  async exchangeRefreshToken(client: OAuthClientInformationFull, refreshToken: string, scopes?: string[]): Promise<OAuthTokens> {
    const rec = await this.store.take<TokenRecord>(`rt:${sha256(refreshToken)}`);
    if (!rec || rec.clientId !== client.client_id) throw new InvalidGrantError("Invalid refresh token");
    if (!(await this.store.get(`grant:${rec.grantId}`))) throw new InvalidGrantError("Microsoft sign-in was revoked; please reconnect");
    return this.issueTokens(rec.grantId, client.client_id, scopes ?? rec.scopes, rec.resource);
  }

  private async issueTokens(grantId: string, clientId: string, scopes: string[], resource?: string): Promise<OAuthTokens> {
    const accessToken = randomToken(32);
    const refreshToken = randomToken(32);
    const expiresAt = Math.floor(Date.now() / 1000) + ACCESS_TOKEN_TTL;
    await this.store.set(`at:${sha256(accessToken)}`, { grantId, clientId, scopes, resource, expiresAt } satisfies TokenRecord, ACCESS_TOKEN_TTL);
    await this.store.set(`rt:${sha256(refreshToken)}`, { grantId, clientId, scopes, resource } satisfies TokenRecord, REFRESH_TOKEN_TTL);
    return {
      access_token: accessToken,
      token_type: "Bearer",
      expires_in: ACCESS_TOKEN_TTL,
      refresh_token: refreshToken,
      scope: scopes.join(" ") || undefined,
    };
  }

  async verifyAccessToken(token: string): Promise<AuthInfo> {
    const rec = await this.store.get<TokenRecord>(`at:${sha256(token)}`);
    if (!rec) throw new InvalidTokenError("Invalid or expired access token");
    if (!(await this.store.get(`grant:${rec.grantId}`))) throw new InvalidTokenError("Microsoft sign-in was revoked; please reconnect");
    return {
      token,
      clientId: rec.clientId,
      scopes: rec.scopes,
      expiresAt: rec.expiresAt,
      resource: rec.resource ? new URL(rec.resource) : undefined,
      extra: { grantId: rec.grantId },
    };
  }

  async revokeToken(_client: OAuthClientInformationFull, request: { token: string }): Promise<void> {
    await this.store.delete(`at:${sha256(request.token)}`);
    await this.store.delete(`rt:${sha256(request.token)}`);
  }

  // ---- Microsoft token management ----
  private async saveGrant(grantId: string, grant: Grant) {
    await this.store.set(`grant:${grantId}`, { enc: encrypt(this.cfg.encryptionKey, JSON.stringify(grant)) }, GRANT_TTL);
  }

  private async loadGrant(grantId: string): Promise<Grant | undefined> {
    const row = await this.store.get<{ enc: string }>(`grant:${grantId}`);
    return row ? (JSON.parse(decrypt(this.cfg.encryptionKey, row.enc)) as Grant) : undefined;
  }

  /** Returns a valid Microsoft Graph access token for this grant, refreshing if needed. */
  async getMicrosoftToken(grantId: string, forceRefresh = false): Promise<string> {
    const grant = await this.loadGrant(grantId);
    if (!grant) throw new ReauthRequiredError("Your Microsoft sign-in has expired. Reconnect the OneNote connector in Claude.");
    if (!forceRefresh && grant.msExpiresAt - Date.now() > 2 * 60 * 1000) return grant.msAccessToken;

    // One refresh per grant at a time (Microsoft rotates refresh tokens).
    const inflight = this.refreshLocks.get(grantId);
    if (inflight) return inflight;
    const p = (async () => {
      try {
        const t = await this.ms.refresh(grant.msRefreshToken);
        await this.saveGrant(grantId, {
          msRefreshToken: t.refresh_token ?? grant.msRefreshToken,
          msAccessToken: t.access_token,
          msExpiresAt: Date.now() + t.expires_in * 1000,
        });
        return t.access_token;
      } catch (err) {
        if (err instanceof MicrosoftOAuthError && ["invalid_grant", "interaction_required", "consent_required"].includes(err.code)) {
          await this.store.delete(`grant:${grantId}`);
          throw new ReauthRequiredError("Your Microsoft sign-in has expired or was revoked. Reconnect the OneNote connector in Claude.");
        }
        throw err;
      } finally {
        this.refreshLocks.delete(grantId);
      }
    })();
    this.refreshLocks.set(grantId, p);
    return p;
  }
}

function withParams(uri: string, params: Record<string, string | undefined>): string {
  const u = new URL(uri);
  for (const [k, v] of Object.entries(params)) if (v !== undefined) u.searchParams.set(k, v);
  return u.toString();
}

function parseCookies(header?: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of (header ?? "").split(";")) {
    const i = part.indexOf("=");
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}

