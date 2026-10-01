import pg from "pg";

/** Minimal JSON key-value store with optional TTL. */
export interface KVStore {
  get<T>(key: string): Promise<T | undefined>;
  set(key: string, value: unknown, ttlSeconds?: number): Promise<void>;
  delete(key: string): Promise<void>;
  /** Atomically read and delete (for one-time auth codes). */
  take<T>(key: string): Promise<T | undefined>;
}

export class MemoryStore implements KVStore {
  private data = new Map<string, { value: unknown; expiresAt?: number }>();

  async get<T>(key: string): Promise<T | undefined> {
    const entry = this.data.get(key);
    if (!entry) return undefined;
    if (entry.expiresAt && entry.expiresAt < Date.now()) {
      this.data.delete(key);
      return undefined;
    }
    return entry.value as T;
  }

  async set(key: string, value: unknown, ttlSeconds?: number): Promise<void> {
    this.data.set(key, { value, expiresAt: ttlSeconds ? Date.now() + ttlSeconds * 1000 : undefined });
  }

  async delete(key: string): Promise<void> {
    this.data.delete(key);
  }

  async take<T>(key: string): Promise<T | undefined> {
    const v = await this.get<T>(key);
    this.data.delete(key);
    return v;
  }
}

export class PostgresStore implements KVStore {
  private pool: pg.Pool;
  private ready: Promise<void>;

  constructor(connectionString: string) {
    this.pool = new pg.Pool({ connectionString, max: 10 });
    this.ready = this.init();
    // Sweep expired rows every 10 minutes.
    setInterval(() => {
      this.pool.query("DELETE FROM kv WHERE expires_at IS NOT NULL AND expires_at < now()").catch(() => {});
    }, 10 * 60 * 1000).unref();
  }

  private async init() {
    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS kv (
        key text PRIMARY KEY,
        value jsonb NOT NULL,
        expires_at timestamptz
      )`);
  }

  async get<T>(key: string): Promise<T | undefined> {
    await this.ready;
    const r = await this.pool.query(
      "SELECT value FROM kv WHERE key = $1 AND (expires_at IS NULL OR expires_at > now())",
      [key],
    );
    return r.rows[0]?.value as T | undefined;
  }

  async set(key: string, value: unknown, ttlSeconds?: number): Promise<void> {
    await this.ready;
    await this.pool.query(
      `INSERT INTO kv (key, value, expires_at) VALUES ($1, $2, $3)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, expires_at = EXCLUDED.expires_at`,
      [key, JSON.stringify(value), ttlSeconds ? new Date(Date.now() + ttlSeconds * 1000) : null],
    );
  }

  async delete(key: string): Promise<void> {
    await this.ready;
    await this.pool.query("DELETE FROM kv WHERE key = $1", [key]);
  }

  async take<T>(key: string): Promise<T | undefined> {
    await this.ready;
    const r = await this.pool.query(
      "DELETE FROM kv WHERE key = $1 AND (expires_at IS NULL OR expires_at > now()) RETURNING value",
      [key],
    );
    return r.rows[0]?.value as T | undefined;
  }
}
