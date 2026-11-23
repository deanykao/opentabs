import {
  ToolError,
  findLocalStorageEntry,
  getPreScriptValue,
  notifyReadinessChanged,
  parseRetryAfterMs,
  sleep,
  waitUntil,
} from '@opentabs-dev/plugin-sdk';
import { findMsalTokens } from './msal.js';

// ---------------------------------------------------------------------------
// Environment detection
// ---------------------------------------------------------------------------

type TeamsEnvironment = 'consumer' | 'enterprise';

/** Detect whether we are running on consumer or enterprise Teams. */
const detectEnvironment = (): TeamsEnvironment => {
  try {
    if (typeof window !== 'undefined' && window.location.hostname === 'teams.live.com') {
      return 'consumer';
    }
  } catch {
    // Fall through to enterprise default
  }
  return 'enterprise';
};

/** Origin of the current Teams tab (enterprise tabs run on teams.microsoft.com or teams.cloud.microsoft). */
const currentOrigin = (): string => {
  try {
    if (typeof window !== 'undefined') return window.location.origin;
  } catch {
    // fall through
  }
  return 'https://teams.microsoft.com';
};

const getAuthzUrl = (): string =>
  detectEnvironment() === 'consumer'
    ? 'https://teams.live.com/api/auth/v1.0/authz/consumer'
    : `${currentOrigin()}/api/authsvc/v1.0/authz`;

const CONSUMER_CHAT_BASE = 'https://teams.live.com/api/chatsvc/consumer';

// ---------------------------------------------------------------------------
// Pre-script values and rejected-token tracking
// ---------------------------------------------------------------------------

interface CapturedToken {
  secret: string;
  expiresOn: number;
}

/**
 * Read a value the pre-script stashed under the `teams` plugin namespace.
 *
 * `getPreScriptValue` from the SDK depends on `globalThis.__openTabs._pluginName`
 * being bound, but the adapter IIFE only binds that during tool dispatch —
 * `isReady()` runs before any tool is dispatched, so the SDK helper would
 * return `undefined` there. We try the SDK helper first, then fall back to a
 * direct namespace read against the documented path.
 */
const readPreScriptValue = <T>(key: string): T | undefined => {
  const viaSdk = getPreScriptValue<T>(key);
  if (viaSdk !== undefined) return viaSdk;
  const ns = (globalThis as { __openTabs?: { preScript?: Record<string, Record<string, unknown>> } }).__openTabs
    ?.preScript?.teams;
  return ns?.[key] as T | undefined;
};

/**
 * Secrets (Skype JWTs or MSAL access tokens) the server answered 401 to.
 * The pre-script keeps handing back the last JWT it captured; without this
 * set a rejected JWT would be reused forever.
 */
const rejectedSecrets = new Set<string>();
const MAX_REJECTED = 32;

const rejectSecret = (secret: string): void => {
  if (rejectedSecrets.size >= MAX_REJECTED) {
    const oldest = rejectedSecrets.values().next().value;
    if (oldest !== undefined) rejectedSecrets.delete(oldest);
  }
  rejectedSecrets.add(secret);
};

const usableCaptured = (captured: CapturedToken | undefined, headroomSec: number): CapturedToken | null => {
  if (!captured || typeof captured.secret !== 'string' || captured.secret.length === 0) return null;
  if (typeof captured.expiresOn !== 'number' || captured.expiresOn <= Date.now() / 1000 + headroomSec) return null;
  if (rejectedSecrets.has(captured.secret)) return null;
  return captured;
};

/**
 * Find an unexpired MSAL access token for the Skype API. The MSAL localStorage
 * cache is read directly (it is plaintext on current Teams v2 builds); the
 * pre-script's captured copy is the fallback for builds where the cache is
 * unreadable at probe time.
 */
const getSkypeAccessToken = (): string | null => {
  const fromCache = findMsalTokens('api.spaces.skype.com', 60).find(t => !rejectedSecrets.has(t.secret));
  if (fromCache) return fromCache.secret;
  const key = detectEnvironment() === 'consumer' ? 'consumerToken' : 'enterpriseToken';
  return usableCaptured(readPreScriptValue<CapturedToken>(key), 60)?.secret ?? null;
};

// ---------------------------------------------------------------------------
// Chat service URL discovery
// ---------------------------------------------------------------------------

/** Where the chat service base URL came from — surfaced in 404 errors. */
type RegionSource = 'consumer' | 'discovered' | 'authz' | 'jwt-region-claim' | 'default-amer';

let cachedChatBase: { base: string; source: RegionSource } | null = null;

/** Decode a JWT payload without verification (we only need to read claims). */
const decodeJwtPayload = (jwt: string): Record<string, unknown> => {
  const parts = jwt.split('.');
  if (parts.length < 2) return {};
  try {
    // base64url payloads are emitted without `=` padding; restore it before
    // atob, which throws on lengths not divisible by 4.
    const raw = parts[1]?.replace(/-/g, '+').replace(/_/g, '/') ?? '';
    const padded = raw + '='.repeat((4 - (raw.length % 4)) % 4);
    return JSON.parse(atob(padded)) as Record<string, unknown>;
  } catch {
    return {};
  }
};

/** Pick chatServiceAfd (same-origin proxy) over chatService from a regionGtms-shaped object. */
const pickChatService = (gtms: { chatService?: string; chatServiceAfd?: string } | undefined): string | null =>
  gtms?.chatServiceAfd || gtms?.chatService || null;

/** Read the chat service URL from the region config the Teams SPA caches in localStorage. */
const discoverChatBaseFromStorage = (): string | null => {
  const skypeToken = findLocalStorageEntry(key => key.includes('Discover.SKYPE-TOKEN'));
  if (skypeToken) {
    try {
      const data = JSON.parse(skypeToken.value) as {
        item?: { regionGtms?: { chatService?: string; chatServiceAfd?: string } };
      };
      const found = pickChatService(data.item?.regionGtms);
      if (found) return found;
    } catch {
      // fall through to the next source
    }
  }
  const regionGtm = findLocalStorageEntry(key => key.includes('Discover.DISCOVER-REGION-GTM'));
  if (regionGtm) {
    try {
      const data = JSON.parse(regionGtm.value) as { item?: { chatService?: string; chatServiceAfd?: string } };
      const found = pickChatService(data.item);
      if (found) return found;
    } catch {
      // fall through
    }
  }
  return null;
};

/** Region claim (`rgn`, e.g. "amer") of the freshest Skype JWT we can see without a network call. */
const regionFromKnownJwt = (): string | null => {
  const jwt = cachedSkypeJwt?.token ?? usableCaptured(readPreScriptValue<CapturedToken>('skypeJwt'), 0)?.secret;
  if (!jwt) return null;
  const rgn = decodeJwtPayload(jwt).rgn;
  return typeof rgn === 'string' && /^[a-z0-9-]{2,16}$/i.test(rgn) ? rgn.toLowerCase() : null;
};

/**
 * Resolve the chat service base URL. Enterprise order: value learned from the
 * authz response, Teams' own discovery cache in localStorage, the JWT `rgn`
 * claim, and only then a hard-coded AMER default.
 */
const resolveChatBase = (): { base: string; source: RegionSource } => {
  if (detectEnvironment() === 'consumer') return { base: CONSUMER_CHAT_BASE, source: 'consumer' };
  if (cachedChatBase) return cachedChatBase;

  const stored = discoverChatBaseFromStorage();
  if (stored) {
    cachedChatBase = { base: stored, source: 'discovered' };
    return cachedChatBase;
  }
  const rgn = regionFromKnownJwt();
  if (rgn) {
    // Not cached: a later authz response or discovery entry is more authoritative.
    return { base: `${currentOrigin()}/api/chatsvc/${rgn}`, source: 'jwt-region-claim' };
  }
  return { base: `${currentOrigin()}/api/chatsvc/amer`, source: 'default-amer' };
};

/** Chat service base URL for the current environment (no trailing slash). */
export const getChatServiceBase = (): string => resolveChatBase().base;

// ---------------------------------------------------------------------------
// Skype JWT acquisition
// ---------------------------------------------------------------------------

interface SkypeIdentity {
  skypeid: string;
  signinname: string;
}

let cachedSkypeJwt: { token: string; expiresAt: number } | null = null;
let cachedIdentity: { identity: SkypeIdentity; expiresAt: number } | null = null;
/** Identity fields from the last authz response (consumer returns them directly). */
let exchangeIdentity: Partial<SkypeIdentity> = {};
let inflightExchange: Promise<string> | null = null;

const JWT_HEADROOM_MS = 60_000;

const cachedJwtIfFresh = (): string | null => {
  if (cachedSkypeJwt && Date.now() < cachedSkypeJwt.expiresAt - JWT_HEADROOM_MS) return cachedSkypeJwt.token;
  return null;
};

/** True when a Skype JWT can be obtained without user interaction. */
export const isTeamsAuthenticated = (): boolean =>
  cachedJwtIfFresh() !== null ||
  usableCaptured(readPreScriptValue<CapturedToken>('skypeJwt'), 60) !== null ||
  getSkypeAccessToken() !== null;

/**
 * Wait for a usable Skype auth path. Bounded to stay well under the 5 s the
 * extension allows `isReady()`.
 */
export const waitForTeamsAuth = (timeoutMs = 3500): Promise<boolean> =>
  waitUntil(() => isTeamsAuthenticated(), { interval: 250, timeout: timeoutMs }).then(
    () => true,
    () => false,
  );

let readinessWatcher: Promise<void> | null = null;

/**
 * When `isReady()` gave up, keep watching for auth in the background and tell the
 * extension to re-probe the moment it appears instead of waiting for the 30 s poll.
 * Skipped outside a page: the build validator calls `isReady()` in Node, where the
 * pending watcher would hold the build open for its full timeout.
 */
export const watchForAuthAndNotify = (): void => {
  if (readinessWatcher || typeof document === 'undefined') return;
  readinessWatcher = waitUntil(() => isTeamsAuthenticated(), { interval: 500, timeout: 120_000 })
    .then(() => notifyReadinessChanged())
    .catch(() => undefined)
    .finally(() => {
      readinessWatcher = null;
    });
};

/** Exchange the MSAL Skype access token for a Skype JWT via authsvc. */
const exchangeSkypeJwt = async (): Promise<string> => {
  const accessToken = getSkypeAccessToken();
  if (!accessToken) {
    throw ToolError.auth(
      `Not authenticated — no Skype API access token found for ${detectEnvironment()} Teams. Open or reload the Teams tab and sign in.`,
    );
  }

  let response: Response;
  try {
    response = await fetch(getAuthzUrl(), {
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
      credentials: 'include',
      signal: AbortSignal.timeout(15_000),
    });
  } catch (err) {
    if (err instanceof DOMException && err.name === 'TimeoutError') {
      throw ToolError.timeout('Skype JWT exchange timed out');
    }
    throw ToolError.internal(`Skype JWT exchange failed: ${err instanceof Error ? err.message : String(err)}`);
  }

  if (!response.ok) {
    const body = await response.text().catch(() => '');
    if (response.status === 401 || response.status === 403) {
      rejectSecret(accessToken);
      cachedSkypeJwt = null;
      throw ToolError.auth(`Skype JWT exchange auth error (${String(response.status)}): ${body.substring(0, 256)}`);
    }
    throw ToolError.internal(`Skype JWT exchange error (${String(response.status)}): ${body.substring(0, 256)}`);
  }

  const data = (await response.json()) as {
    // Consumer: { skypeToken: { skypetoken, expiresIn, skypeid, signinname } }
    skypeToken?: { skypetoken?: string; expiresIn?: number; skypeid?: string; signinname?: string };
    // Enterprise: { tokens: { skypeToken, expiresIn }, regionGtms: { chatService, chatServiceAfd } }
    tokens?: { skypeToken?: string; expiresIn?: number };
    regionGtms?: { chatService?: string; chatServiceAfd?: string };
  };

  const jwt = data.tokens?.skypeToken ?? data.skypeToken?.skypetoken;
  if (!jwt) throw ToolError.internal('Skype JWT exchange returned empty token');

  const expiresIn = data.tokens?.expiresIn ?? data.skypeToken?.expiresIn ?? 3600;
  cachedSkypeJwt = { token: jwt, expiresAt: Date.now() + expiresIn * 1000 };

  if (data.skypeToken?.skypeid) {
    exchangeIdentity = { skypeid: data.skypeToken.skypeid, signinname: data.skypeToken.signinname ?? '' };
  }
  if (detectEnvironment() === 'enterprise') {
    const base = pickChatService(data.regionGtms);
    if (base) cachedChatBase = { base, source: 'authz' };
  }

  notifyReadinessChanged();
  return jwt;
};

/**
 * Get a Skype JWT. Sources, in priority order:
 *
 * 1. Adapter-local cache.
 * 2. `skypeJwt` pre-script slot (Teams' own authsvc response), unless the server
 *    already rejected that exact JWT.
 * 3. Our own authsvc exchange using the MSAL Skype access token.
 */
const getSkypeJwt = async (): Promise<string> => {
  const cached = cachedJwtIfFresh();
  if (cached) return cached;

  const preCapture = usableCaptured(readPreScriptValue<CapturedToken>('skypeJwt'), 60);
  if (preCapture) {
    cachedSkypeJwt = { token: preCapture.secret, expiresAt: preCapture.expiresOn * 1000 };
    return preCapture.secret;
  }

  inflightExchange ??= exchangeSkypeJwt().finally(() => {
    inflightExchange = null;
  });
  return inflightExchange;
};

/** Forget a JWT the server rejected so it is neither cached nor re-read from the pre-script. */
const invalidateSkypeJwt = (jwt: string): void => {
  rejectSecret(jwt);
  if (cachedSkypeJwt?.token === jwt) cachedSkypeJwt = null;
  cachedIdentity = null;
};

// ---------------------------------------------------------------------------
// Current user identity
// ---------------------------------------------------------------------------

/**
 * The pre-script decodes the MSAL ID token and stashes
 * `preferred_username` / `upn` / `email` under `signInName`.
 */
const findSignInName = (): string => readPreScriptValue<string>('signInName') ?? '';

/**
 * Current user's Skype identity (MRI suffix and sign-in email). Read from the
 * Skype JWT claims; cached for the JWT lifetime.
 */
export const getSkypeIdentity = async (): Promise<SkypeIdentity> => {
  if (cachedIdentity && Date.now() < cachedIdentity.expiresAt) return cachedIdentity.identity;

  let jwt = await getSkypeJwt();
  let claims = decodeJwtPayload(jwt);
  if (typeof claims.skypeid !== 'string' && !exchangeIdentity.skypeid) {
    // Neither the JWT nor a previous authz response named the user: force an exchange to learn it.
    jwt = await exchangeSkypeJwt();
    claims = decodeJwtPayload(jwt);
  }

  const skypeid = typeof claims.skypeid === 'string' ? claims.skypeid : (exchangeIdentity.skypeid ?? '');
  const identity: SkypeIdentity = { skypeid, signinname: findSignInName() || exchangeIdentity.signinname || '' };
  const jwtExpMs = typeof claims.exp === 'number' ? claims.exp * 1000 : (cachedSkypeJwt?.expiresAt ?? Date.now());
  if (skypeid) cachedIdentity = { identity, expiresAt: jwtExpMs - JWT_HEADROOM_MS };
  return identity;
};

// ---------------------------------------------------------------------------
// Shared request pipeline (401 re-auth, bounded 429 retry, error classification)
// ---------------------------------------------------------------------------

const MAX_RETRY_AFTER_MS = 5000;
const DEFAULT_RETRY_AFTER_MS = 1000;

interface Authorizer {
  /** Fetch credentials mode: Skype endpoints use cookies, bearer APIs answer with ACAO:* and so must be sent without. */
  credentials: RequestCredentials;
  /** Headers for the next attempt (may fetch/exchange tokens). */
  headers(): Promise<Record<string, string>>;
  /** Called on 401: evict the credential used by the last `headers()` call. */
  reject(): void;
}

interface RequestSpec {
  url: string;
  method: string;
  body?: unknown;
  /** Short label used in error messages. */
  label: string;
  /** Extra detail appended to not-found errors. */
  notFoundHint?: () => string;
}

const skypeAuthorizer = (): Authorizer => {
  let used = '';
  return {
    credentials: 'include',
    headers: async () => {
      used = await getSkypeJwt();
      return { Authentication: `skypetoken=${used}`, 'Content-Type': 'application/json' };
    },
    reject: () => {
      if (used) invalidateSkypeJwt(used);
    },
  };
};

const bearerAuthorizer = (host: string, scope?: string): Authorizer => {
  let used = '';
  return {
    credentials: 'omit',
    headers: () => {
      const token = findMsalTokens(host, 30, scope).find(t => !rejectedSecrets.has(t.secret));
      if (!token) {
        return Promise.reject(
          ToolError.auth(
            `No unexpired ${host} access token found in the Teams MSAL cache${scope ? ` (wanted scope ${scope})` : ''}. ` +
              'The cache may be encrypted or Teams has not acquired a token for this service yet — open Teams in the tab, use the relevant feature once, and retry.',
          ),
        );
      }
      used = token.secret;
      return Promise.resolve({ Authorization: `Bearer ${used}` });
    },
    reject: () => {
      if (used) rejectSecret(used);
    },
  };
};

/** Pull the human-readable message out of a Teams/Graph error body. */
const extractServerMessage = (text: string): string => {
  const trimmed = text.trim();
  if (trimmed.startsWith('{')) {
    try {
      const j = JSON.parse(trimmed) as { error?: { message?: string } | string; message?: string; errorCode?: string };
      const msg =
        (typeof j.error === 'object' ? j.error?.message : typeof j.error === 'string' ? j.error : undefined) ??
        j.message ??
        j.errorCode;
      if (msg) return msg.substring(0, 512);
    } catch {
      // fall through to raw text
    }
  }
  return trimmed.substring(0, 512);
};

/** Classify HTTP errors into ToolError categories. Always throws. */
const classifyHttpError = (response: Response, errorText: string, spec: RequestSpec): never => {
  const { status } = response;
  const where = `${spec.method} ${spec.label}`;
  const message = extractServerMessage(errorText);

  if (status === 429) {
    const header = response.headers.get('Retry-After');
    throw ToolError.rateLimited(
      `Teams API rate limited: ${where} — ${message}`,
      header !== null ? parseRetryAfterMs(header) : undefined,
    );
  }
  if (status === 401) {
    throw ToolError.auth(`Teams API auth error (401): ${where} — ${message}. Reload the Teams tab and retry.`);
  }
  if (status === 403) {
    throw new ToolError(
      `Teams API permission denied (403): ${where} — ${message || 'the signed-in user or token lacks permission for this operation'}`,
      'PERMISSION_DENIED',
      { category: 'validation', retryable: false },
    );
  }
  if (status === 404) {
    throw ToolError.notFound(`Teams API not found: ${where} — ${message}${spec.notFoundHint?.() ?? ''}`);
  }
  if (status === 400) {
    throw ToolError.validation(`Teams API bad request: ${where} — ${message}`);
  }
  throw ToolError.internal(`Teams API error (${String(status)}): ${where} — ${message}`);
};

/**
 * Send a request through the shared policy: on 401 evict the credential that
 * was just rejected and retry once with a fresh one; on 429 retry once when the
 * server's Retry-After is at most 5 s; everything else is classified and thrown.
 */
const send = async (az: Authorizer, spec: RequestSpec): Promise<Response> => {
  let retried401 = false;
  let retried429 = false;

  for (;;) {
    const headers = await az.headers();
    let response: Response;
    try {
      response = await fetch(spec.url, {
        method: spec.method,
        // A Content-Type on body-less requests would force a CORS preflight.
        headers: spec.body === undefined ? headers : { 'Content-Type': 'application/json', ...headers },
        body: spec.body === undefined ? undefined : JSON.stringify(spec.body),
        credentials: az.credentials,
        signal: AbortSignal.timeout(30_000),
      });
    } catch (err) {
      if (err instanceof DOMException && err.name === 'TimeoutError') {
        throw ToolError.timeout(`Teams API timed out: ${spec.method} ${spec.label}`);
      }
      throw ToolError.internal(
        `Network error calling Teams API: ${spec.method} ${spec.label} — ${err instanceof Error ? err.message : String(err)}`,
      );
    }

    if (response.ok) return response;

    if (response.status === 401 && !retried401) {
      retried401 = true;
      az.reject();
      continue;
    }
    if (response.status === 429 && !retried429) {
      const header = response.headers.get('Retry-After');
      const waitMs = header !== null ? parseRetryAfterMs(header) : DEFAULT_RETRY_AFTER_MS;
      if (waitMs !== undefined && waitMs <= MAX_RETRY_AFTER_MS) {
        retried429 = true;
        await sleep(waitMs);
        continue;
      }
    }

    const errorText = await response.text().catch(() => '');
    return classifyHttpError(response, errorText, spec);
  }
};

/** Parse a response body as JSON; empty bodies (201/204) yield `{}`. */
const readJson = async <T>(response: Response, spec: RequestSpec): Promise<T> => {
  const text = await response.text();
  if (text.length === 0) return {} as T;
  try {
    return JSON.parse(text) as T;
  } catch {
    throw ToolError.internal(`Teams API returned invalid JSON: ${spec.method} ${spec.label}`);
  }
};

type Query = Record<string, string | number | boolean | undefined>;

const withQuery = (url: string, query?: Query): string => {
  if (!query) return url;
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(query)) {
    if (value !== undefined) params.append(key, String(value));
  }
  const qs = params.toString();
  if (!qs) return url;
  return url + (url.includes('?') ? '&' : '?') + qs;
};

// ---------------------------------------------------------------------------
// Chat Service API (Skype-JWT authenticated)
// ---------------------------------------------------------------------------

interface ChatRequestOptions {
  method?: string;
  body?: unknown;
  query?: Query;
}

/**
 * Authenticated request to the Teams Chat Service. `endpoint` starts with `/v1/`.
 * Returns the parsed body together with the response headers (Location etc.).
 */
export const chatApiRaw = async <T>(
  endpoint: string,
  options: ChatRequestOptions = {},
): Promise<{ data: T; headers: Headers }> => {
  const { base, source } = resolveChatBase();
  const spec: RequestSpec = {
    url: withQuery(`${base}${endpoint}`, options.query),
    method: options.method ?? 'GET',
    body: options.body,
    label: endpoint,
    notFoundHint: () =>
      source === 'default-amer'
        ? ` (no region could be discovered, so the default AMER chat service ${base} was used — the account may live in another region)`
        : ` (chat service: ${base})`,
  };
  const response = await send(skypeAuthorizer(), spec);
  return { data: await readJson<T>(response, spec), headers: response.headers };
};

export const chatApi = async <T>(endpoint: string, options: ChatRequestOptions = {}): Promise<T> =>
  (await chatApiRaw<T>(endpoint, options)).data;

/** Create a chat thread. Returns the thread ID extracted from the Location response header. */
export const createThread = async (
  members: Array<{ id: string; role: string }>,
  properties?: Record<string, unknown>,
): Promise<string> => {
  const body: Record<string, unknown> = { members };
  if (properties) body.properties = properties;
  const { headers } = await chatApiRaw<Record<string, unknown>>('/v1/threads', { method: 'POST', body });
  const threadId = (headers.get('Location') ?? '').split('/threads/').pop() ?? '';
  if (!threadId) throw ToolError.internal('Thread created but no thread ID returned in Location header');
  return threadId;
};

/** Request against `/v1/threads/{threadId}{subpath}` (member management, properties). */
export const threadApi = <T>(
  threadId: string,
  subpath: string,
  options: { method?: string; body?: unknown } = {},
): Promise<T> => chatApi<T>(`/v1/threads/${encodeURIComponent(threadId)}${subpath}`, options);

// ---------------------------------------------------------------------------
// MSAL-bearer APIs (Graph, presence, Substrate search)
// ---------------------------------------------------------------------------

export const GRAPH_HOST = 'graph.microsoft.com';
const GRAPH_BASE = `https://${GRAPH_HOST}/v1.0`;

interface BearerOptions {
  method?: string;
  body?: unknown;
  query?: Query;
  /** Scope name a token should carry (preferred, not required). */
  scope?: string;
}

const bearerApi = async <T>(host: string, url: string, label: string, options: BearerOptions): Promise<T> => {
  const spec: RequestSpec = {
    url: withQuery(url, options.query),
    method: options.method ?? 'GET',
    body: options.body,
    label,
  };
  const response = await send(bearerAuthorizer(host, options.scope), spec);
  return readJson<T>(response, spec);
};

/** Microsoft Graph request. `path` starts with `/` (relative to /v1.0). */
export const graphApi = <T>(path: string, options: BearerOptions = {}): Promise<T> =>
  bearerApi<T>(GRAPH_HOST, `${GRAPH_BASE}${path}`, `graph ${path}`, options);

/** Follow a Graph `@odata.nextLink`. Rejects links that do not point at Graph so the bearer token cannot be sent elsewhere. */
export const graphApiFromLink = <T>(link: string, options: BearerOptions = {}): Promise<T> => {
  let parsed: URL;
  try {
    parsed = new URL(link);
  } catch {
    throw ToolError.validation('Invalid cursor: not a URL');
  }
  if (parsed.protocol !== 'https:' || parsed.hostname !== GRAPH_HOST) {
    throw ToolError.validation('Invalid cursor: it must be a next_cursor value returned by this tool');
  }
  return bearerApi<T>(GRAPH_HOST, link, 'graph (next page)', options);
};

export const presenceApi = <T>(path: string, options: BearerOptions = {}): Promise<T> =>
  bearerApi<T>(
    'presence.teams.microsoft.com',
    `https://presence.teams.microsoft.com${path}`,
    `presence ${path}`,
    options,
  );

export const substrateApi = <T>(path: string, options: BearerOptions = {}): Promise<T> =>
  bearerApi<T>('substrate.office.com', `https://substrate.office.com${path}`, `substrate ${path}`, options);

/** Clear all module-level caches so the next API call re-discovers fresh values. */
export const clearCaches = (): void => {
  cachedChatBase = null;
  cachedSkypeJwt = null;
  cachedIdentity = null;
  exchangeIdentity = {};
  inflightExchange = null;
  rejectedSecrets.clear();
};
