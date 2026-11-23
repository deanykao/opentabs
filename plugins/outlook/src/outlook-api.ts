import {
  ToolError,
  buildQueryString,
  clearAuthCache,
  getAuthCache,
  getLocalStorage,
  setAuthCache,
  waitUntil,
} from '@opentabs-dev/plugin-sdk';
import type { FetchFromPageOptions } from '@opentabs-dev/plugin-sdk';

const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';
const OUTLOOK_API_BASE = 'https://outlook.office.com/api/v2.0';
const OUTLOOK_API_BETA_BASE = 'https://outlook.office.com/api/beta';

// Outlook enterprise MSAL client ID
const MSAL_CLIENT_ID = '9199bf20-a13f-4107-85dc-02114787ef48';
// Consumer fallback
const MSAL_CLIENT_ID_CONSUMER = '2821b473-fe24-4c86-ba16-62834d6e80c3';

interface OutlookAuth {
  token: string;
  apiBase: string; // which API base URL this token works with
  /**
   * Lowercased scope names the token grants on its API host (e.g. `calendars.readwrite`).
   * Undefined when the cache entry does not expose its scopes (MSAL v1).
   */
  scopes?: string[];
}

/**
 * A request capability. Mail and calendar can require different Graph scopes, and
 * a single token is not guaranteed to carry both — enterprise tenants commonly
 * issue a narrowly-scoped Graph token (User.* only) alongside a broad Outlook REST
 * token. Calendar capabilities rank candidates by their scope claim: tokens that
 * grant the capability are tried first, tokens with unknown scopes (`.default`,
 * MSAL v1) after, and tokens whose scopes provably lack it are skipped. A 403 from
 * a token that grants the capability is a genuine permission denial and is surfaced
 * as such instead of cascading. Each capability caches its winning token in its own
 * slot so mail, calendar-read, and calendar-write never evict each other.
 */
type Capability = 'mail' | 'calendar' | 'calendar-write';

/** Per-capability auth cache key, keeping each token bucket separate. */
const AUTH_CACHE_KEY: Record<Capability, string> = {
  mail: 'outlook',
  calendar: 'outlook-calendar',
  'calendar-write': 'outlook-calendar-write',
};

/** Scope a token must carry for a capability; `null` means no scope filtering. */
const CAPABILITY_SCOPE: Record<Capability, RegExp | null> = {
  mail: null,
  calendar: /^calendars\./,
  'calendar-write': /^calendars\.readwrite/,
};

type ScopeFit = 'grants' | 'unknown' | 'lacks';

export const scopeFit = (auth: Pick<OutlookAuth, 'scopes'>, capability: Capability): ScopeFit => {
  const required = CAPABILITY_SCOPE[capability];
  if (!required) return 'unknown';
  if (!auth.scopes) return 'unknown';
  if (auth.scopes.some(s => required.test(s))) return 'grants';
  return auth.scopes.includes('.default') ? 'unknown' : 'lacks';
};

/** Order candidates for a capability: granting tokens, then unknown; drop those that lack it. */
export const rankCandidates = <A extends Pick<OutlookAuth, 'scopes'>>(candidates: A[], capability: Capability): A[] => {
  const grants = candidates.filter(c => scopeFit(c, capability) === 'grants');
  const unknown = candidates.filter(c => scopeFit(c, capability) === 'unknown');
  return [...grants, ...unknown];
};

/**
 * Endpoint addressing. A plain string is appended to whichever API base the token
 * serves (Graph and Outlook REST share most paths) or, when absolute, is only sent
 * with a token for that same base (OData `@odata.nextLink`). An object gives per-API
 * paths for operations whose routes differ; a token for an API with no path is skipped.
 */
export type ApiEndpoint = string | { graph?: string; rest?: string; restVersion?: 'v2.0' | 'beta' };

const resolveBase = (auth: OutlookAuth, endpoint: ApiEndpoint): { base: string; path: string } | null => {
  if (typeof endpoint === 'string') {
    if (!endpoint.startsWith('https://')) return { base: auth.apiBase, path: endpoint };
    return endpoint.startsWith(`${auth.apiBase}/`) ? { base: '', path: endpoint } : null;
  }
  if (auth.apiBase === GRAPH_API_BASE) return endpoint.graph ? { base: GRAPH_API_BASE, path: endpoint.graph } : null;
  if (!endpoint.rest) return null;
  return { base: endpoint.restVersion === 'beta' ? OUTLOOK_API_BETA_BASE : OUTLOOK_API_BASE, path: endpoint.rest };
};

/**
 * Enumerate every MSAL client id whose token-index key starts with `prefix`.
 * The SDK's `findLocalStorageEntry` returns only the first match, which silently
 * drops every additional client id present when a user has multiple Microsoft
 * apps signed in.
 */
const findAllMsalClientIds = (prefix: string): string[] => {
  const ids: string[] = [];
  try {
    const storage = window.localStorage;
    for (let i = 0; i < storage.length; i++) {
      const key = storage.key(i);
      if (key?.startsWith(prefix)) {
        ids.push(key.slice(prefix.length));
      }
    }
  } catch {
    // SecurityError or missing localStorage — nothing we can do
  }
  return ids;
};

/**
 * Scope names in the AAD scope claim (space-separated) whose URL hostname equals
 * `host`, lowercased and stripped of leading slashes (`https://h//Calendars.Read` →
 * `calendars.read`). Scopes are parsed as URLs and compared by `hostname` rather than
 * substring-matched so a malicious value like `https://attacker.com/graph.microsoft.com/foo`
 * cannot satisfy the check. An empty result means the claim grants nothing on `host`.
 */
export const scopesForHost = (target: string, host: string): string[] => {
  const scopes: string[] = [];
  for (const scope of target.split(/\s+/)) {
    if (scope.length === 0) continue;
    try {
      const url = new URL(scope);
      if (url.hostname.toLowerCase() === host) scopes.push(url.pathname.replace(/^\/+/, '').toLowerCase());
    } catch {
      // non-URL scopes (openid, profile, email, ...) — skip
    }
  }
  return scopes;
};

/**
 * Return unexpired access tokens whose target scope claim grants the given host
 * from the MSAL v2 or v3 cache. Both versions share the same entry shape
 * (`secret`, `target`, `expiresOn`); only the index-key prefix differs.
 */
const findMsalModernTokens = (version: '2' | '3', clientId: string, host: string): OutlookAuth[] => {
  const tokenKeysRaw = getLocalStorage(`msal.${version}.token.keys.${clientId}`);
  if (!tokenKeysRaw) return [];

  let tokenKeys: { accessToken?: string[] };
  try {
    tokenKeys = JSON.parse(tokenKeysRaw);
  } catch {
    return [];
  }
  if (!tokenKeys.accessToken) return [];

  const apiBase = host === 'graph.microsoft.com' ? GRAPH_API_BASE : OUTLOOK_API_BASE;
  const matches: OutlookAuth[] = [];
  for (const key of tokenKeys.accessToken) {
    const raw = getLocalStorage(key);
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (typeof parsed.secret !== 'string' || parsed.secret.length === 0) continue;

      const scopes = scopesForHost(parsed.target ?? '', host);
      if (scopes.length === 0) continue;

      // Strict numeric coercion — Number.parseInt would accept '9999999999junk'
      // as a giant future expiry; Number(...) returns NaN for trailing garbage.
      const expiresOn = Number(parsed.expiresOn);
      if (!Number.isInteger(expiresOn) || expiresOn <= 0 || expiresOn * 1000 < Date.now()) continue;

      matches.push({ token: parsed.secret, apiBase, scopes });
    } catch {
      // skip invalid entries
    }
  }
  return matches;
};

/** Search MSAL v1 token cache for valid Graph API access tokens. */
const findMsalV1Tokens = (clientId: string): OutlookAuth[] => {
  const tokenKeysRaw = getLocalStorage(`msal.token.keys.${clientId}`);
  if (!tokenKeysRaw) return [];

  let tokenKeys: { accessToken?: string[] };
  try {
    tokenKeys = JSON.parse(tokenKeysRaw);
  } catch {
    return [];
  }
  if (!tokenKeys.accessToken) return [];

  const matches: OutlookAuth[] = [];
  for (const key of tokenKeys.accessToken) {
    if (!/(?:^|[\s/])graph\.microsoft\.com(?:[/\s]|$)/.test(key)) continue;
    const raw = getLocalStorage(key);
    if (!raw) continue;
    try {
      const parsed = JSON.parse(raw);
      if (typeof parsed.secret !== 'string' || parsed.secret.length === 0) continue;
      const expiresOn = Number(parsed.expiresOn);
      if (!Number.isInteger(expiresOn) || expiresOn <= 0 || expiresOn * 1000 < Date.now()) continue;
      matches.push({ token: parsed.secret, apiBase: GRAPH_API_BASE });
    } catch {
      // skip invalid entries
    }
  }
  return matches;
};

/**
 * Return every MSAL-cached token plausibly usable against Graph or Outlook REST,
 * deduplicated and ordered by preference (v3 enterprise → v2 enterprise → v1
 * consumer → other client ids). `api()` cascades through the list on 401/403;
 * the first token the API accepts is cached for subsequent calls.
 */
const collectAuthCandidates = (): OutlookAuth[] => {
  const all: OutlookAuth[] = [];

  // Enterprise, known client id — v3 then v2, Graph then Outlook REST per version
  for (const version of ['3', '2'] as const) {
    all.push(...findMsalModernTokens(version, MSAL_CLIENT_ID, 'graph.microsoft.com'));
    all.push(...findMsalModernTokens(version, MSAL_CLIENT_ID, 'outlook.office.com'));
  }

  // Consumer v1
  all.push(...findMsalV1Tokens(MSAL_CLIENT_ID_CONSUMER));

  // Fallback: every other client id present in localStorage, modern then v1.
  // Enumerate (not first-match) so users with multiple Microsoft apps signed in
  // surface every token, not just the first index key the iterator hits.
  for (const version of ['3', '2'] as const) {
    for (const cid of findAllMsalClientIds(`msal.${version}.token.keys.`)) {
      if (cid === MSAL_CLIENT_ID) continue;
      all.push(...findMsalModernTokens(version, cid, 'graph.microsoft.com'));
      all.push(...findMsalModernTokens(version, cid, 'outlook.office.com'));
    }
  }
  for (const cid of findAllMsalClientIds('msal.token.keys.')) {
    if (cid === MSAL_CLIENT_ID_CONSUMER) continue;
    all.push(...findMsalV1Tokens(cid));
  }

  // Deduplicate by token — multiple lookups can surface the same secret
  const seen = new Set<string>();
  return all.filter(c => {
    if (seen.has(c.token)) return false;
    seen.add(c.token);
    return true;
  });
};

export const isAuthenticated = (): boolean => {
  if (getAuthCache<OutlookAuth>(AUTH_CACHE_KEY.mail)) return true;
  return collectAuthCandidates().length > 0;
};

export const waitForAuth = (): Promise<boolean> =>
  waitUntil(() => isAuthenticated(), { interval: 500, timeout: 5000 }).then(
    () => true,
    () => false,
  );

/**
 * Recursively convert PascalCase keys to camelCase.
 * Outlook REST API returns PascalCase; Graph returns camelCase.
 * Normalizing to camelCase lets all mappers work with both APIs.
 */
const toCamelCase = (str: string): string => str.charAt(0).toLowerCase() + str.slice(1);

const normalizeKeys = (obj: unknown): unknown => {
  if (obj === null || obj === undefined || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(normalizeKeys);
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(obj as Record<string, unknown>)) {
    // Skip OData metadata keys like @odata.context
    const newKey = key.startsWith('@') ? key : toCamelCase(key);
    result[newKey] = normalizeKeys(value);
  }
  return result;
};

const toPascalCase = (str: string): string => str.charAt(0).toUpperCase() + str.slice(1);

/**
 * Outlook REST property names that are not the PascalCase of their Graph name.
 * `timeConstraint.timeSlots` is `TimeConstraint.Timeslots` on REST v2.0.
 */
const REST_KEY_OVERRIDES: Record<string, string> = { timeSlots: 'Timeslots' };

/**
 * Recursively convert object keys to PascalCase.
 * The Outlook REST API deserializes request bodies case-sensitively and requires
 * PascalCase property names for both entities (Subject, Start/DateTime) and OData
 * action parameters (Schedules, Comment). Graph uses camelCase. Request bodies are
 * authored in camelCase and transformed here when the request targets the REST base;
 * GET query options ($filter, $orderby) are case-insensitive on REST and unaffected.
 * Enum values are left as authored — REST accepts Graph's camelCase enum values.
 */
export const pascalCaseKeys = (obj: unknown): unknown => {
  if (obj === null || obj === undefined || typeof obj !== 'object') return obj;
  if (Array.isArray(obj)) return obj.map(pascalCaseKeys);
  const result: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(obj as Record<string, unknown>)) {
    result[REST_KEY_OVERRIDES[key] ?? toPascalCase(key)] = pascalCaseKeys(value);
  }
  return result;
};

interface RequestOptions {
  method?: string;
  body?: unknown;
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
}

type RequestOutcome<T> =
  | { kind: 'ok'; data: T }
  /** The token's API has no route for this endpoint — try the next candidate. */
  | { kind: 'unsupported' }
  /** 401/403 — the token was rejected; `message` is the server's error text. */
  | { kind: 'rejected'; status: number; message: string };

const readErrorMessage = async (response: Response): Promise<string | undefined> => {
  try {
    const errBody = (await response.json()) as { error?: { message?: string } };
    return errBody.error?.message;
  } catch {
    return undefined;
  }
};

/**
 * Send an authenticated request and handle the response.
 * Returns the parsed response, `unsupported` when the token's API has no route for
 * the endpoint, or `rejected` on 401/403 so the caller can decide whether to cascade.
 * Throws on every other error.
 */
const sendRequest = async <T>(
  auth: OutlookAuth,
  endpoint: ApiEndpoint,
  options: RequestOptions,
): Promise<RequestOutcome<T>> => {
  const target = resolveBase(auth, endpoint);
  if (!target) return { kind: 'unsupported' };
  const isOutlookApi = auth.apiBase === OUTLOOK_API_BASE;

  // Outlook REST API uses different $select field names, so drop $select
  // and let it return all fields. The normalizeKeys step handles casing.
  const query = options.query ? { ...options.query } : undefined;
  if (isOutlookApi && query) {
    delete (query as Record<string, unknown>).$select;
  }

  const qs = query ? buildQueryString(query) : '';
  const url = qs ? `${target.base}${target.path}?${qs}` : `${target.base}${target.path}`;
  const method = options.method ?? 'GET';

  const headers: Record<string, string> = {
    Authorization: `Bearer ${auth.token}`,
    ...options.headers,
  };

  const init: FetchFromPageOptions = { method, headers };

  if (options.body) {
    headers['Content-Type'] = 'application/json';
    const body = isOutlookApi ? pascalCaseKeys(options.body) : options.body;
    init.body = JSON.stringify(body);
  }

  let response: Response;
  try {
    response = await fetch(url, {
      ...init,
      credentials: 'omit',
      signal: AbortSignal.timeout(30_000),
    });
  } catch (err) {
    if (err instanceof DOMException && err.name === 'TimeoutError') {
      throw ToolError.timeout('Microsoft API request timed out.');
    }
    if (err instanceof DOMException && err.name === 'AbortError') {
      throw new ToolError('Request aborted', 'aborted');
    }
    throw new ToolError(`Network error: ${err instanceof Error ? err.message : 'unknown'}`, 'network_error', {
      category: 'internal',
      retryable: true,
    });
  }

  if (response.status === 204) return { kind: 'ok', data: {} as T };

  if (response.status === 429) {
    const retryAfter = response.headers.get('Retry-After');
    const retryMs = retryAfter ? Number.parseInt(retryAfter, 10) * 1000 : undefined;
    throw ToolError.rateLimited('Microsoft API rate limit exceeded.', retryMs);
  }

  if (response.status === 401 || response.status === 403) {
    const message = (await readErrorMessage(response)) ?? `Microsoft API error (${response.status})`;
    return { kind: 'rejected', status: response.status, message };
  }

  if (response.status === 404) {
    throw ToolError.notFound('The requested resource was not found.');
  }

  if (!response.ok) {
    const errorMsg = (await readErrorMessage(response)) ?? `Microsoft API error (${response.status})`;
    if (response.status === 400 || response.status === 422) {
      throw ToolError.validation(errorMsg);
    }
    throw ToolError.internal(errorMsg);
  }

  // Successful actions (cancel, RSVP, sendMail) often return 202/205 or a 200 with
  // an empty body and no JSON. Only parse when the response actually carries JSON.
  const contentType = response.headers.get('content-type') ?? '';
  if (!contentType.includes('application/json')) return { kind: 'ok', data: {} as T };

  const json = await response.json();
  return { kind: 'ok', data: (isOutlookApi ? normalizeKeys(json) : json) as T };
};

/**
 * Make an authenticated request to a Microsoft 365 API for the given capability,
 * trying the capability's cached winner first, then every MSAL-cached candidate
 * ranked by scope (see `Capability`), and caching the first that succeeds under the
 * capability's slot. Automatically uses whichever API the resolved token supports
 * (Graph or Outlook REST). A 403 from a token whose scopes grant the capability is
 * reported as a permission denial with the server's message; otherwise 401/403
 * cascade to the next candidate.
 */
export const api = async <T>(
  endpoint: ApiEndpoint,
  options: RequestOptions = {},
  capability: Capability = 'mail',
): Promise<T> => {
  const cacheKey = AUTH_CACHE_KEY[capability];
  let lastRejection: { status: number; message: string } | undefined;

  const attempt = async (auth: OutlookAuth): Promise<{ data: T } | null> => {
    const outcome = await sendRequest<T>(auth, endpoint, options);
    if (outcome.kind === 'ok') return { data: outcome.data };
    if (outcome.kind === 'rejected') {
      if (outcome.status === 403 && scopeFit(auth, capability) === 'grants') {
        throw ToolError.auth(`Microsoft 365 denied this request: ${outcome.message}`, 'PERMISSION_DENIED');
      }
      lastRejection = outcome;
    }
    return null;
  };

  const cached = getAuthCache<OutlookAuth>(cacheKey);
  if (cached) {
    const r = await attempt(cached);
    if (r) return r.data;
    if (lastRejection) clearAuthCache(cacheKey);
  }

  const candidates = rankCandidates(collectAuthCandidates(), capability);
  // Skip the cached candidate we just tried.
  const remaining = cached ? candidates.filter(c => c.token !== cached.token) : candidates;
  if (remaining.length === 0 && !lastRejection) {
    throw ToolError.auth('Not authenticated — please sign in to Microsoft 365.');
  }

  for (const auth of remaining) {
    const r = await attempt(auth);
    if (r) {
      setAuthCache(cacheKey, auth);
      return r.data;
    }
  }

  if (lastRejection?.status === 403) {
    throw ToolError.auth(`Microsoft 365 denied this request: ${lastRejection.message}`, 'PERMISSION_DENIED');
  }
  if (lastRejection) throw ToolError.auth('Authentication expired — please refresh the Outlook page.');
  throw ToolError.validation('No signed-in Microsoft 365 API supports this operation for your account.');
};

interface ODataPage<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

/**
 * Fetch an OData collection, following `@odata.nextLink` until `limit` items are
 * collected or the collection ends. Next links are absolute URLs on the API that
 * served the first page; `api()` only sends them with a token for that same base.
 */
export const apiPaged = async <T>(
  endpoint: ApiEndpoint,
  options: RequestOptions,
  capability: Capability,
  limit: number,
): Promise<{ items: T[]; hasMore: boolean }> => {
  const items: T[] = [];
  let page = await api<ODataPage<T>>(endpoint, options, capability);
  for (;;) {
    for (const item of page.value ?? []) {
      if (items.length === limit) return { items, hasMore: true };
      items.push(item);
    }
    const next = page['@odata.nextLink'];
    if (!next) return { items, hasMore: false };
    if (items.length === limit) return { items, hasMore: true };
    page = await api<ODataPage<T>>(next, { headers: options.headers }, capability);
  }
};

// OWS gateway lives on the OWA page's own origin — a third base distinct from Graph
// and Outlook REST. It serves the client-side compose settings (roaming signatures,
// startup data) that Graph's mailboxSettings does not expose. The adapter runs on
// whichever OWA host matched (outlook.cloud.microsoft, outlook.office.com, or
// outlook.office365.com), so resolve against the current origin to stay same-origin
// rather than hardcoding one host. Read lazily (not at module scope) because the
// plugin module is also loaded in Node at build time, where `window` is undefined.
const owsBaseUrl = (): string => window.location.origin;
const OWS_AUTH_CACHE_KEY = 'outlook-ows';

/** MSAL access-token claims OWS routing headers are derived from. */
interface OwsTokenClaims {
  puid?: string;
  tid?: string;
}

/**
 * Read the `puid`/`tid` claims from an MSAL access token's JWT payload. No signature
 * verification — the token is already trusted (we minted the request with it); we
 * only need its routing hints. Returns an empty object for any malformed token.
 */
const decodeTokenClaims = (jwt: string): OwsTokenClaims => {
  const payload = jwt.split('.')[1];
  if (!payload) return {};
  try {
    const json = atob(payload.replace(/-/g, '+').replace(/_/g, '/'));
    const parsed = JSON.parse(json) as Record<string, unknown>;
    return {
      puid: typeof parsed.puid === 'string' ? parsed.puid : undefined,
      tid: typeof parsed.tid === 'string' ? parsed.tid : undefined,
    };
  } catch {
    return {};
  }
};

/**
 * Build the header set OWS gateway endpoints expect. The mailbox anchor
 * (`x-anchormailbox` / `x-routingparameter-sessionkey`) is derived from the token's
 * own `puid`/`tid` claims so the gateway routes to the right mailbox's settings.
 */
const buildOwsHeaders = (token: string, extra?: Record<string, string>): Record<string, string> => {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/json',
    'x-ms-appname': 'owa-reactmail',
    owaappid: '00000002-0000-0ff1-ce00-000000000000',
    'x-outlook-client': 'owa',
    ...extra,
  };
  const { puid, tid } = decodeTokenClaims(token);
  if (puid) {
    const anchor = tid ? `PUID:${puid}@${tid}` : `PUID:${puid}`;
    headers['x-anchormailbox'] = anchor;
    headers['x-routingparameter-sessionkey'] = anchor;
  }
  return headers;
};

/**
 * Encode an OWS query string with `encodeURIComponent` per value, yielding `%20` for
 * spaces (in signature display names) and `%2C` for commas (the `settingname` list
 * delimiter) — both of which the OWS gateway accepts. `URLSearchParams` is avoided
 * because it encodes spaces as `+`, which OWS does not decode back to a space.
 */
const encodeOwsQuery = (query: Record<string, string | number | boolean | undefined>): string => {
  const parts: string[] = [];
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined) continue;
    parts.push(`${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
  }
  return parts.join('&');
};

interface OwsRequestOptions {
  method?: string;
  query?: Record<string, string | number | boolean | undefined>;
  headers?: Record<string, string>;
}

type OwsOutcome<T> = { kind: 'ok'; data: T } | { kind: 'notFound' } | { kind: 'authFail' };

/**
 * Send one OWS request with a specific token. OWS is same-origin, so cookies ride
 * along with the Bearer token. A 404 means the token authenticated but the settings
 * collection holds no such item (e.g. a signature with no body) — distinct from
 * 401/403, which signals "try the next candidate token".
 */
const sendOwsRequest = async <T>(
  token: string,
  endpoint: string,
  options: OwsRequestOptions,
): Promise<OwsOutcome<T>> => {
  const qs = options.query ? encodeOwsQuery(options.query) : '';
  const base = owsBaseUrl();
  const url = qs ? `${base}${endpoint}?${qs}` : `${base}${endpoint}`;

  let response: Response;
  try {
    response = await fetch(url, {
      method: options.method ?? 'GET',
      headers: buildOwsHeaders(token, options.headers),
      credentials: 'same-origin',
      signal: AbortSignal.timeout(30_000),
    });
  } catch (err) {
    if (err instanceof DOMException && err.name === 'TimeoutError') {
      throw ToolError.timeout('Microsoft settings request timed out.');
    }
    throw new ToolError(`Network error: ${err instanceof Error ? err.message : 'unknown'}`, 'network_error', {
      category: 'internal',
      retryable: true,
    });
  }

  if (response.status === 401 || response.status === 403) return { kind: 'authFail' };
  if (response.status === 404) return { kind: 'notFound' };
  if (response.status === 429) {
    const retryAfter = response.headers.get('Retry-After');
    const retryMs = retryAfter ? Number.parseInt(retryAfter, 10) * 1000 : undefined;
    throw ToolError.rateLimited('Microsoft API rate limit exceeded.', retryMs);
  }
  if (!response.ok) {
    throw ToolError.internal(`Microsoft settings request failed (${response.status}).`);
  }

  const contentType = response.headers.get('content-type') ?? '';
  const data = contentType.includes('application/json') ? ((await response.json()) as T) : ({} as T);
  return { kind: 'ok', data };
};

/**
 * Request an OWS gateway endpoint on the OWA origin (roaming signature settings,
 * startup data), cascading through every MSAL candidate on 401/403 exactly like
 * `api()` and caching the winner in its own slot. Returns `undefined` when a token
 * authenticated but the endpoint returned 404 (no such setting), so callers can
 * treat a missing signature as "none configured" rather than an error. Throws only
 * when no candidate authenticates at all.
 */
export const owsRequest = async <T>(endpoint: string, options: OwsRequestOptions = {}): Promise<T | undefined> => {
  const cached = getAuthCache<OutlookAuth>(OWS_AUTH_CACHE_KEY);
  if (cached) {
    const outcome = await sendOwsRequest<T>(cached.token, endpoint, options);
    if (outcome.kind === 'ok') return outcome.data;
    if (outcome.kind === 'notFound') return undefined;
    clearAuthCache(OWS_AUTH_CACHE_KEY);
  }

  const candidates = collectAuthCandidates();
  const remaining = cached ? candidates.filter(c => c.token !== cached.token) : candidates;
  if (remaining.length === 0) {
    throw ToolError.auth(
      cached
        ? 'Authentication expired — please refresh the Outlook page.'
        : 'Not authenticated — please sign in to Microsoft 365.',
    );
  }

  for (const auth of remaining) {
    const outcome = await sendOwsRequest<T>(auth.token, endpoint, options);
    if (outcome.kind === 'ok') {
      setAuthCache(OWS_AUTH_CACHE_KEY, auth);
      return outcome.data;
    }
    // A 404 means this token authenticated (the gateway resolved its mailbox) and the
    // setting is simply absent — a valid winner to cache, so a genuinely-missing
    // signature does not re-cascade through every token on each subsequent call.
    if (outcome.kind === 'notFound') {
      setAuthCache(OWS_AUTH_CACHE_KEY, auth);
      return undefined;
    }
  }

  throw ToolError.auth('Authentication expired — please refresh the Outlook page.');
};
