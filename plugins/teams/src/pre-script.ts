import { definePreScript } from '@opentabs-dev/plugin-sdk/pre-script';

/**
 * Pre-script for the Teams plugin.
 *
 * Runs at document_start in MAIN world, strictly before any page script.
 * Two capture paths run in parallel:
 *
 * 1. **MSAL localStorage observer** — scans existing localStorage entries and
 *    hooks `Storage.prototype.setItem` to catch future writes. Recognises MSAL
 *    credential entries by their *value* shape (`credentialType`, `target`,
 *    `secret`, `expiresOn`) and stashes the Skype API access token and the
 *    sign-in name from the ID token. The adapter also reads the MSAL cache
 *    directly (it is plaintext on current Teams builds); this capture is the
 *    fallback for the moments before the cache is populated.
 *
 * 2. **authsvc fetch interceptor** — wraps `window.fetch` to intercept Teams'
 *    own POST to `/api/authsvc/v1.0/authz` (and the consumer equivalent). When
 *    Teams exchanges its MSAL Skype token for a Skype JWT, we clone the
 *    response body and stash the resulting JWT under `skypeJwt`.
 *
 * Adapter reads via `getPreScriptValue` keys:
 *   - `consumerToken`  → { secret, expiresOn } MSAL Skype token for teams.live.com
 *   - `enterpriseToken`→ { secret, expiresOn } MSAL Skype token for enterprise Teams
 *   - `skypeJwt`       → { secret, expiresOn } Skype JWT captured from authsvc response
 *   - `signInName`     → preferred_username / upn / email from the ID token
 */
definePreScript(({ set, log }) => {
  const SKYPE_HOST_PATTERN = /spaces\.skype\.com/i;
  const AUTHSVC_PATTERN = /\/authsvc\/v1\.0\/authz|\/auth\/v1\.0\/authz/;

  type CapturedToken = { secret: string; expiresOn: number };

  const decodeJwtClaims = (jwt: string): Record<string, unknown> | null => {
    const parts = jwt.split('.');
    if (parts.length < 2) return null;
    try {
      // base64url payloads are emitted without `=` padding; restore it before
      // calling atob, which throws InvalidCharacterError on lengths not
      // divisible by 4.
      const raw = (parts[1] ?? '').replace(/-/g, '+').replace(/_/g, '/');
      const padded = raw + '='.repeat((4 - (raw.length % 4)) % 4);
      return JSON.parse(atob(padded)) as Record<string, unknown>;
    } catch {
      return null;
    }
  };

  // ---------------------------------------------------------------------------
  // Path 1: MSAL localStorage observer
  // ---------------------------------------------------------------------------

  /**
   * Inspect a single localStorage value. If it parses as an MSAL credential
   * cache entry of interest, stash the relevant pieces.
   */
  const inspect = (value: string): void => {
    if (typeof value !== 'string' || value.length < 32) return;
    let parsed: unknown;
    try {
      parsed = JSON.parse(value);
    } catch {
      return;
    }
    // MSAL credential entries are JSON objects. Anything else (null, primitives,
    // arrays) is unrelated cache content; bail before property access so a stray
    // `null` value can't throw and abort the initial-scan loop.
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) return;
    const entry = parsed as Record<string, unknown>;

    const credentialType = String(entry.credentialType ?? '').toLowerCase();
    const secret = typeof entry.secret === 'string' ? entry.secret : '';
    if (!secret) return;

    if (credentialType === 'accesstoken') {
      const target = String(entry.target ?? '');
      if (!SKYPE_HOST_PATTERN.test(target)) return;

      const expiresOn = Number.parseInt(String(entry.expiresOn ?? '0'), 10);
      if (!Number.isFinite(expiresOn) || expiresOn <= Date.now() / 1000) return;

      const captured: CapturedToken = { secret, expiresOn };
      // Route by current page hostname so the slot we write matches the
      // adapter's detectEnvironment() (which also keys off hostname).
      const slot = window.location.hostname === 'teams.live.com' ? 'consumerToken' : 'enterpriseToken';
      set(slot, captured);
      log.debug(`[teams] captured ${slot}`);
      return;
    }

    if (credentialType === 'idtoken') {
      const claims = decodeJwtClaims(secret);
      if (!claims) return;
      const signInName = String(claims.preferred_username ?? claims.upn ?? claims.email ?? '');
      if (signInName) {
        set('signInName', signInName);
        // Don't log the value — preferred_username / upn / email is PII.
        log.debug('[teams] captured sign-in name from ID token');
      }
    }
  };

  // ---------------------------------------------------------------------------
  // Path 2: authsvc fetch interceptor (Skype JWT)
  // ---------------------------------------------------------------------------

  /**
   * Wrap window.fetch to intercept Teams' own POST to the authsvc endpoint.
   * When Teams exchanges its encrypted MSAL Skype token for a Skype JWT during
   * startup, we clone the successful response body and stash the JWT under
   * `skypeJwt`. This fires for every authsvc call (initial + token refreshes),
   * keeping the cached JWT fresh without any separate polling.
   *
   * The interceptor runs in MAIN world at document_start, before any Teams JS,
   * so it is guaranteed to be in place for the startup authsvc call.
   */
  const FETCH_MARKER = '__opentabsTeamsFetchPatched';
  const winAny = window as unknown as Record<string, unknown>;
  if (!winAny[FETCH_MARKER]) {
    const realFetch = window.fetch.bind(window);
    window.fetch = function patchedFetch(
      input: Parameters<typeof fetch>[0],
      init?: Parameters<typeof fetch>[1],
    ): ReturnType<typeof fetch> {
      const url = input instanceof Request ? input.url : String(input);
      if (AUTHSVC_PATTERN.test(url)) {
        return realFetch(input, init).then(response => {
          if (response.ok) {
            response
              .clone()
              .json()
              .then((data: unknown) => {
                if (!data || typeof data !== 'object') return;
                const d = data as Record<string, unknown>;

                // Enterprise format: { tokens: { skypeToken: string, expiresIn: number } }
                const tokens = d.tokens as Record<string, unknown> | undefined;
                const entJwt = typeof tokens?.skypeToken === 'string' ? tokens.skypeToken : null;
                if (entJwt) {
                  const expiresIn = typeof tokens?.expiresIn === 'number' ? tokens.expiresIn : 3600;
                  set('skypeJwt', {
                    secret: entJwt,
                    expiresOn: Math.floor(Date.now() / 1000) + expiresIn,
                  } as CapturedToken);
                  log.debug('[teams] captured Skype JWT from authsvc response');
                  return;
                }

                // Consumer format: { skypeToken: { skypetoken: string, expiresIn: number } }
                const skypeToken = d.skypeToken as Record<string, unknown> | undefined;
                const conJwt = typeof skypeToken?.skypetoken === 'string' ? skypeToken.skypetoken : null;
                if (conJwt) {
                  const expiresIn = typeof skypeToken?.expiresIn === 'number' ? skypeToken.expiresIn : 3600;
                  set('skypeJwt', {
                    secret: conJwt,
                    expiresOn: Math.floor(Date.now() / 1000) + expiresIn,
                  } as CapturedToken);
                  log.debug('[teams] captured Skype JWT from authsvc response (consumer)');
                }
              })
              .catch(() => {
                // Never block the page on JSON parse failure.
              });
          }
          return response;
        });
      }
      return realFetch(input, init);
    };
    winAny[FETCH_MARKER] = true;
  }

  // ---------------------------------------------------------------------------
  // Initial scans
  // ---------------------------------------------------------------------------

  // localStorage: MSAL credential entries (Path 1).
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key) continue;
      const val = localStorage.getItem(key);
      if (val) inspect(val);
    }
  } catch {
    // localStorage access can throw under sandboxing; ignore.
  }

  // ---------------------------------------------------------------------------
  // setItem hook (Path 1)
  // ---------------------------------------------------------------------------

  // Idempotent: if the prototype is already patched (re-injection during
  // dev hot reload, future iframe-realm reuse, etc.), skip rather than
  // stack wrappers.
  const PATCH_MARKER = '__opentabsTeamsSetItemPatched';
  const proto = Storage.prototype as Storage & Record<string, unknown>;
  if (!proto[PATCH_MARKER]) {
    const realSetItem = Storage.prototype.setItem;
    Storage.prototype.setItem = function patchedSetItem(key: string, value: string) {
      try {
        if (this === localStorage) inspect(value);
      } catch {
        // Never break the page if our observer throws.
      }
      return realSetItem.call(this, key, value);
    };
    proto[PATCH_MARKER] = true;
  }

  log.info('[teams] MSAL cache observer + authsvc interceptor installed');
});
