## Exploration: Fix persistent OAuth authentication (token refresh)

### Current State

The MCP server requires re-authentication every few hours because the OAuth token refresh mechanism is broken. There are **three separate token systems** that don't communicate, with **four different scope lists** across the codebase.

### Affected Areas

- `outlook-auth-server.js` — Standalone OAuth server (port 3333). Handles initial code exchange. Saves tokens to `~/.outlook-mcp-tokens.json`. **Scopes (lines 46-54):** `offline_access`, `User.Read`, `Mail.Read`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`. **Missing:** `Mail.ReadWrite`, `Files.Read`, `Files.ReadWrite`.
- `auth/token-storage.js` — TokenStorage class used by the MCP server at runtime. Has `refreshAccessToken()` (lines 120-197) and `exchangeCodeForTokens()` (lines 200-265). **Default scopes (line 20):** `offline_access User.Read Mail.Read` — only 3 scopes. **Missing:** `Mail.Send`, `Mail.ReadWrite`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Files.Read`, `Files.ReadWrite`.
- `auth/token-manager.js` — Legacy token manager. Only loads tokens and checks expiration. **No refresh logic.** Reads from `config.AUTH_CONFIG.tokenStorePath`.
- `auth/index.js` — Creates a singleton `TokenStorage` instance. `ensureAuthenticated()` calls `tokenStorage.getValidAccessToken()` which triggers refresh.
- `config.js` — `AUTH_CONFIG.scopes` (line 23): `['Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'User.Read', 'Calendars.Read', 'Calendars.ReadWrite', 'Files.Read', 'Files.ReadWrite']`. **Missing:** `offline_access`, `Contacts.Read`. **This config is NEVER USED** for actual token operations — it's only referenced by `token-manager.js` (which has no refresh) and `auth/tools.js` (which only generates the auth URL).
- `auth/oauth-server.js` — Alternative Express-based OAuth module. Uses `createAuthConfig()` which defaults scopes to `'offline_access User.Read Mail.Read'` (same broken default as token-storage). **Not actively used** — the standalone `outlook-auth-server.js` is the one in use.
- `power-automate/` — Uses separate Flow API tokens (`flow_access_token`, `flow_refresh_token`) stored alongside Graph tokens in the same file. Managed by `token-manager.js` `saveFlowTokens()`.

### Scope Inconsistency Map

| Source | Scopes | Missing vs. Full Set |
|--------|--------|---------------------|
| **Full set needed** | `offline_access`, `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Files.Read`, `Files.ReadWrite` | — |
| `outlook-auth-server.js` (lines 46-54) | `offline_access`, `User.Read`, `Mail.Read`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read` | ❌ `Mail.ReadWrite`, `Files.Read`, `Files.ReadWrite` |
| `auth/token-storage.js` (line 20, default) | `offline_access`, `User.Read`, `Mail.Read` | ❌ `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Files.Read`, `Files.ReadWrite` |
| `config.js` (line 23) | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`, `Calendars.Read`, `Calendars.ReadWrite`, `Files.Read`, `Files.ReadWrite` | ❌ `offline_access`, `Contacts.Read` |
| `auth/oauth-server.js` (default) | `offline_access`, `User.Read`, `Mail.Read` | ❌ `Mail.ReadWrite`, `Mail.Send`, `Calendars.Read`, `Calendars.ReadWrite`, `Contacts.Read`, `Files.Read`, `Files.ReadWrite` |

### Root Cause Analysis

**Primary cause: Scope mismatch between auth server and token-storage.**

1. **Auth server (`outlook-auth-server.js`)** requests scopes including `offline_access` during initial code exchange. Microsoft returns an `access_token` + `refresh_token` and saves them to `~/.outlook-mcp-tokens.json`. This works fine.

2. **Token-storage (`auth/token-storage.js`)** has its OWN default scopes: `'offline_access User.Read Mail.Read'` (line 20). When `refreshAccessToken()` fires (line 137), it sends `scope: this.config.scopes.join(' ')` — which is only those 3 scopes. **Microsoft may return a downscoped token** because the refresh request asks for fewer permissions than the original token.

3. **The `config.js` scopes are unused** for actual token operations. `token-manager.js` reads `config.AUTH_CONFIG.tokenStorePath` for the file path, but never calls refresh. `token-storage.js` has its own hardcoded default scopes and its own `tokenStorePath` — it does NOT read from `config.js`.

4. **The `MS_SCOPES` env var** (`token-storage.js` line 20) is the escape hatch: if set, it overrides the default. But it's NOT set in `.env` and NOT documented anywhere.

**Secondary issues:**

5. **`outlook-auth-server.js` is also missing scopes** — it doesn't include `Mail.ReadWrite` or `Files.Read`/`Files.ReadWrite`. So even the initial token may lack permissions for some operations.

6. **`auth/oauth-server.js`** (the Express module) has the SAME broken default scopes as `token-storage.js`. It's not actively used (the standalone server is), but it's a trap for anyone who tries to use it.

7. **`token-manager.js`** has NO refresh logic at all. It's used by `auth/tools.js` for `check-auth-status` and `authenticate` handlers. When `check-auth-status` runs, it loads tokens via `token-manager.loadTokenCache()` which checks expiration — if expired, it returns `null` (not authenticated). But `token-manager` never refreshes, so the user sees "Not authenticated" even though a valid `refresh_token` exists.

8. **The refresh_token IS saved** by the auth server (line 307: `fs.writeFileSync(AUTH_CONFIG.tokenStorePath, ...)`). The token response from Microsoft includes `refresh_token` when `offline_access` is in the scope. So the refresh_token IS in the file. The problem is that `token-storage` uses the wrong scopes when it tries to refresh.

9. **The auth server does NOT need to stay running** for refresh to work. Refresh happens entirely in the MCP server process via `token-storage.refreshAccessToken()`. The auth server is only needed for the initial code exchange.

10. **No race condition protection** — `token-storage.js` has `_refreshPromise` dedup (line 126-129) which prevents concurrent refresh attempts. This is good.

### Approaches

1. **Unify all scope lists to a single source of truth** — Define scopes in ONE place (e.g., `config.js`), and have both `outlook-auth-server.js` and `token-storage.js` read from it.
   - Pros: Single point of truth; all systems stay in sync; minimal code changes
   - Cons: `config.js` currently missing `offline_access` and `Contacts.Read` — must add them
   - Effort: **Low**

2. **Fix `token-storage.js` default scopes** — Change line 20 to include the full scope set. Also fix `outlook-auth-server.js` to add missing scopes.
   - Pros: Simplest change; no refactoring needed
   - Cons: Still two separate scope lists that could drift; `MS_SCOPES` env var override could still cause issues
   - Effort: **Low**

3. **Make `token-storage.js` read scopes from `config.js`** — Remove the hardcoded default in `token-storage.js` and have it use `config.AUTH_CONFIG.scopes` (after adding `offline_access` to it).
   - Pros: Single source of truth; `config.js` already has the most complete scope list (just missing `offline_access` and `Contacts.Read`)
   - Cons: `token-storage.js` currently ignores `config.js` entirely; need to add the import
   - Effort: **Low**

4. **Replace `token-manager.js` with `token-storage.js` everywhere** — `token-manager.js` is used by `auth/tools.js` for `check-auth-status` and `authenticate`. These should use `token-storage` instead so they benefit from refresh.
   - Pros: Eliminates the stale token-manager path; `check-auth-status` would report accurate status
   - Cons: More changes; `token-manager.js` also handles Flow tokens which `token-storage` doesn't
   - Effort: **Medium**

### Recommendation

**Approach 1 + 4 combined: Unify scopes in `config.js` AND migrate `token-manager.js` consumers to `token-storage.js`.**

Rationale:
- The scope mismatch is the PRIMARY root cause. Fixing it in one place prevents future drift.
- `config.js` is the natural home for configuration — it already has the most complete scope list.
- `token-manager.js` is dead code for refresh purposes. Its consumers (`check-auth-status`, `authenticate`) should use `token-storage` which actually refreshes.
- Flow tokens are a separate concern and can remain in `token-manager.js` for now.

Concrete changes needed:
1. **`config.js`**: Add `offline_access` and `Contacts.Read` to `AUTH_CONFIG.scopes`
2. **`outlook-auth-server.js`**: Import scopes from `config.js` instead of its own hardcoded list (or at minimum add `Mail.ReadWrite`, `Files.Read`, `Files.ReadWrite`)
3. **`auth/token-storage.js`**: Remove hardcoded default scopes; import from `config.js` (or use `MS_SCOPES` env var as override)
4. **`auth/tools.js`**: `handleCheckAuthStatus` should use `token-storage` instead of `token-manager` so it can report accurate refresh-based status
5. **`auth/oauth-server.js`**: Same scope fix (or document as deprecated)

### Risks

- **Token downgrade on refresh**: If the fix isn't applied before the current token expires, the next refresh will request the wrong scopes and get a downscoped token. Mitigation: force re-auth after the fix is deployed.
- **`offline_access` missing from `config.js`**: If we add it to `config.js` but forget to include it in the scope string sent to Microsoft, the refresh token won't be issued. Mitigation: verify `offline_access` is in the final scope list.
- **`MS_SCOPES` env var override**: If someone has `MS_SCOPES` set in their environment, it will override the config.js value in `token-storage.js`. Mitigation: document this or remove the env var override.
- **Flow tokens**: `token-manager.js` handles Flow tokens separately. Don't break that path.
- **Existing token file**: If the user already has a token file with a downscoped refresh token, even after the fix they may need to re-authenticate. Mitigation: document that re-auth may be needed.

### Ready for Proposal

**Yes.** The root cause is clear — scope mismatch between the auth server and the token refresh mechanism. The fix is well-understood and low-effort. Proceed with `sdd-propose` for the `fix-persistent-auth` change.

Key points to tell the user:
- The primary fix is unifying scope lists so refresh requests the same permissions as the original token
- `config.js` should become the single source of truth for scopes
- `token-manager.js` consumers should migrate to `token-storage.js` for accurate auth status
- A re-authentication may be needed after the fix if the current token was already downscoped
- The auth server does NOT need to stay running for refresh to work
