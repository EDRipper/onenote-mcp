import type { Config } from "./config.js";

export interface MsTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
  id_token?: string;
}

export class MicrosoftOAuthError extends Error {
  constructor(public code: string, description: string) {
    super(description || code);
  }
}

/** Thin wrapper around the Microsoft identity platform v2.0 endpoints. */
export class MicrosoftOAuth {
  constructor(private cfg: Config) {}

  get redirectUri(): string {
    return new URL("/oauth/microsoft/callback", this.cfg.baseUrl).toString();
  }

  private endpoint(path: "authorize" | "token"): string {
    return `${this.cfg.msAuthority}/${this.cfg.msTenant}/oauth2/v2.0/${path}`;
  }

  authorizeUrl(state: string, codeChallenge: string): string {
    const u = new URL(this.endpoint("authorize"));
    u.searchParams.set("client_id", this.cfg.msClientId);
    u.searchParams.set("response_type", "code");
    u.searchParams.set("redirect_uri", this.redirectUri);
    u.searchParams.set("response_mode", "query");
    u.searchParams.set("scope", this.cfg.msScopes.join(" "));
    u.searchParams.set("state", state);
    u.searchParams.set("code_challenge", codeChallenge);
    u.searchParams.set("code_challenge_method", "S256");
    u.searchParams.set("prompt", "select_account");
    return u.toString();
  }

  private async tokenRequest(params: Record<string, string>): Promise<MsTokenResponse> {
    const body = new URLSearchParams({
      client_id: this.cfg.msClientId,
      client_secret: this.cfg.msClientSecret,
      scope: this.cfg.msScopes.join(" "),
      ...params,
    });
    const res = await fetch(this.endpoint("token"), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    });
    const json = (await res.json().catch(() => ({}))) as Record<string, unknown>;
    if (!res.ok || typeof json.access_token !== "string") {
      throw new MicrosoftOAuthError(String(json.error ?? res.status), String(json.error_description ?? ""));
    }
    return json as unknown as MsTokenResponse;
  }

  exchangeCode(code: string, codeVerifier: string): Promise<MsTokenResponse> {
    return this.tokenRequest({
      grant_type: "authorization_code",
      code,
      redirect_uri: this.redirectUri,
      code_verifier: codeVerifier,
    });
  }

  refresh(refreshToken: string): Promise<MsTokenResponse> {
    return this.tokenRequest({ grant_type: "refresh_token", refresh_token: refreshToken });
  }
}
