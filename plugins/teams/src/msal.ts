import { getLocalStorage } from '@opentabs-dev/plugin-sdk';

/** An unexpired MSAL access token read from the Teams localStorage cache. */
export interface MsalToken {
  secret: string;
  /** Expiry as epoch seconds. */
  expiresOn: number;
  /** Space-separated scope claim the token was issued for. */
  target: string;
}

/** Shape of a raw MSAL access-token cache entry (only the fields we read). */
export interface MsalCacheEntry {
  secret?: unknown;
  target?: unknown;
  expiresOn?: unknown;
}

/**
 * True when the AAD scope claim (space-separated) contains at least one scope
 * whose URL hostname equals `host`. Scopes are parsed as URLs and compared by
 * hostname so a value like `https://attacker.com/graph.microsoft.com/foo`
 * cannot satisfy the check.
 */
export const scopeClaimHasHost = (target: string, host: string): boolean => {
  for (const scope of target.split(/\s+/)) {
    if (scope.length === 0) continue;
    try {
      if (new URL(scope).hostname.toLowerCase() === host) return true;
    } catch {
      // non-URL scopes (openid, profile, email, ...) — skip
    }
  }
  return false;
};

/** True when `target` grants the named scope (e.g. `ChannelMessage.Read.All`) on `host`. */
const scopeClaimHasScope = (target: string, host: string, scopeName: string): boolean => {
  const wanted = scopeName.toLowerCase();
  for (const scope of target.split(/\s+/)) {
    try {
      const url = new URL(scope);
      if (url.hostname.toLowerCase() !== host) continue;
      if (url.pathname.split('/').pop()?.toLowerCase() === wanted) return true;
    } catch {
      // non-URL scope — skip
    }
  }
  return false;
};

/**
 * Select the unexpired entries whose scope grants `host`. Tokens that also
 * carry `preferScope` sort first; ties break longest-lived first.
 * `minRemainingSec` demands headroom so a token that expires mid-request is
 * not picked.
 */
export const selectTokens = (
  entries: readonly MsalCacheEntry[],
  host: string,
  nowMs: number,
  minRemainingSec = 0,
  preferScope?: string,
): MsalToken[] => {
  const matches: Array<MsalToken & { preferred: boolean }> = [];
  for (const entry of entries) {
    if (typeof entry.secret !== 'string' || entry.secret.length === 0) continue;
    if (typeof entry.target !== 'string' || !scopeClaimHasHost(entry.target, host)) continue;
    // Strict numeric coercion — parseInt would accept '9999999999junk'.
    const expiresOn = Number(entry.expiresOn);
    if (!Number.isInteger(expiresOn) || expiresOn <= 0) continue;
    if (expiresOn * 1000 <= nowMs + minRemainingSec * 1000) continue;
    matches.push({
      secret: entry.secret,
      expiresOn,
      target: entry.target,
      preferred: preferScope ? scopeClaimHasScope(entry.target, host, preferScope) : false,
    });
  }
  matches.sort((a, b) => Number(b.preferred) - Number(a.preferred) || b.expiresOn - a.expiresOn);
  return matches.map(({ secret, expiresOn, target }) => ({ secret, expiresOn, target }));
};

/** Enumerate every MSAL client id whose token-index key starts with `prefix`. */
const findAllMsalClientIds = (prefix: string): string[] => {
  const ids: string[] = [];
  try {
    const storage = window.localStorage;
    for (let i = 0; i < storage.length; i++) {
      const key = storage.key(i);
      if (key?.startsWith(prefix)) ids.push(key.slice(prefix.length));
    }
  } catch {
    // SecurityError or missing localStorage — nothing we can do
  }
  return ids;
};

/** Read the raw access-token entries listed in the MSAL v2/v3 token index of a client id. */
const readClientEntries = (version: '2' | '3', clientId: string): MsalCacheEntry[] => {
  const indexRaw = getLocalStorage(`msal.${version}.token.keys.${clientId}`);
  if (!indexRaw) return [];
  let keys: string[];
  try {
    const parsed = JSON.parse(indexRaw) as { accessToken?: unknown };
    keys = Array.isArray(parsed.accessToken)
      ? parsed.accessToken.filter((k): k is string => typeof k === 'string')
      : [];
  } catch {
    return [];
  }
  const entries: MsalCacheEntry[] = [];
  for (const key of keys) {
    const raw = getLocalStorage(key);
    if (!raw) continue;
    try {
      const parsed: unknown = JSON.parse(raw);
      if (parsed !== null && typeof parsed === 'object') entries.push(parsed as MsalCacheEntry);
    } catch {
      // Encrypted or malformed entry — skip
    }
  }
  return entries;
};

/**
 * Find unexpired MSAL access tokens granting `host` (e.g. `graph.microsoft.com`),
 * across every MSAL client id and cache version present in localStorage,
 * longest-lived first. Returns an empty list when the cache is absent or encrypted.
 */
export const findMsalTokens = (host: string, minRemainingSec = 0, preferScope?: string): MsalToken[] => {
  const entries: MsalCacheEntry[] = [];
  for (const version of ['2', '3'] as const) {
    for (const clientId of findAllMsalClientIds(`msal.${version}.token.keys.`)) {
      entries.push(...readClientEntries(version, clientId));
    }
  }
  return selectTokens(entries, host, Date.now(), minRemainingSec, preferScope);
};
