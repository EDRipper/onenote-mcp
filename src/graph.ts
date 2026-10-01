export class GraphError extends Error {
  constructor(public status: number, public code: string, message: string) {
    super(message);
  }
}

export interface RequestOptions {
  query?: Record<string, string | number | undefined>;
  body?: unknown;
  /** Defaults to application/json when body is an object. */
  contentType?: string;
  /** Return the Response instead of parsing JSON. */
  raw?: boolean;
}

type TokenGetter = (forceRefresh?: boolean) => Promise<string>;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Tiny Microsoft Graph client scoped to one user's OneNote. */
export class Graph {
  constructor(private base: string, private getToken: TokenGetter) {}

  /** Resolve a path ("/me/onenote/...") or an absolute Graph URL (e.g. @odata.nextLink). */
  private url(pathOrUrl: string, query?: RequestOptions["query"]): string {
    const u = new URL(/^https?:/.test(pathOrUrl) ? pathOrUrl : this.base + pathOrUrl);
    for (const [k, v] of Object.entries(query ?? {})) if (v !== undefined) u.searchParams.set(k, String(v));
    return u.toString();
  }

  async request<T = any>(method: string, path: string, opts: RequestOptions = {}): Promise<T> {
    let forceRefresh = false;
    for (let attempt = 0; ; attempt++) {
      const token = await this.getToken(forceRefresh);
      const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
      let body: string | FormData | undefined;
      if (opts.body !== undefined) {
        if (opts.body instanceof FormData) {
          body = opts.body; // fetch sets multipart/form-data with boundary
        } else if (typeof opts.body === "string") {
          body = opts.body;
          headers["Content-Type"] = opts.contentType ?? "text/html";
        } else {
          body = JSON.stringify(opts.body);
          headers["Content-Type"] = opts.contentType ?? "application/json";
        }
      }
      const res = await fetch(this.url(path, opts.query), { method, headers, body });

      if (res.status === 401 && !forceRefresh) {
        forceRefresh = true;
        continue;
      }
      if ((res.status === 429 || res.status === 503 || res.status === 504) && attempt < 3) {
        const retryAfter = Number(res.headers.get("Retry-After"));
        await sleep(Math.min(Number.isFinite(retryAfter) && retryAfter >= 0 && res.headers.has("Retry-After") ? retryAfter * 1000 : 1000 * 2 ** attempt, 15000));
        continue;
      }
      if (!res.ok) throw await toGraphError(res);
      if (opts.raw) return res as unknown as T;
      if (res.status === 204) return undefined as T;
      const text = await res.text();
      return (text ? JSON.parse(text) : undefined) as T;
    }
  }

  get<T = any>(path: string, query?: RequestOptions["query"]) {
    return this.request<T>("GET", path, { query });
  }

  /** Follow @odata.nextLink until `limit` items are collected. */
  async list<T = any>(path: string, query: RequestOptions["query"] = {}, limit = 500): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = this.url(path, query);
    while (next && out.length < limit) {
      const page: { value: T[]; "@odata.nextLink"?: string } = await this.request("GET", next);
      out.push(...page.value);
      next = page["@odata.nextLink"];
    }
    return out.slice(0, limit);
  }
}

async function toGraphError(res: Response): Promise<GraphError> {
  let code = String(res.status);
  let message = res.statusText;
  try {
    const j: any = await res.json();
    code = j?.error?.code ?? code;
    message = j?.error?.message ?? message;
  } catch {
    /* non-JSON error body */
  }
  const hints: Record<number, string> = {
    403: " (Your account may not allow this. School/work accounts sometimes need IT to approve OneNote access.)",
    404: " (Check the ID — it may have been deleted or moved. List the parent again to get fresh IDs.)",
  };
  return new GraphError(res.status, code, `Microsoft Graph ${res.status} ${code}: ${message}${hints[res.status] ?? ""}`);
}
