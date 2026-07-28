# Tasks: Fix Persistent OAuth Authentication

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~185 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | auto-chain |
| Chain strategy | size-exception |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Focused test command | Runtime harness | Rollback boundary |
|------|------|-----------|----------------------|-----------------|-------------------|
| 1 | Scope unification + refresh fix + auth status | Single PR | `npm test` | N/A — all changes are unit-testable; no runtime auth flow needed | Revert config.js, token-storage.js, outlook-auth-server.js, tools.js, test files |

## Phase 1: RED — Write Failing Tests

- [x] 1.1 `test/auth/token-storage.test.js`: Add test — default scopes come from `config.AUTH_CONFIG.scopes` when no `MS_SCOPES` env var
- [x] 1.2 `test/auth/token-storage.test.js`: Add test — `MS_SCOPES` override without `offline_access` triggers `console.warn`
- [x] 1.3 `test/auth/token-storage.test.js`: Add test — `refreshAccessToken()` POST body `scope` equals `config.AUTH_CONFIG.scopes.join(' ')`
- [x] 1.4 `test/auth/token-storage.test.js`: Add test — `exchangeCodeForTokens()` POST body `scope` equals `config.AUTH_CONFIG.scopes.join(' ')`
- [x] 1.5 `test/auth/tools.test.js`: Create — `handleCheckAuthStatus` returns "Authenticated and ready" when `getValidAccessToken()` returns a token
- [x] 1.6 `test/auth/tools.test.js`: Create — `handleCheckAuthStatus` returns "Not authenticated" when `getValidAccessToken()` returns null

## Phase 2: GREEN — Make Tests Pass

- [x] 2.1 `config.js`: Add `'offline_access'` and `'Contacts.Read'` to `AUTH_CONFIG.scopes` (10 total)
- [x] 2.2 `auth/token-storage.js`: Import `config` from `'../config'`; replace hardcoded `'offline_access User.Read Mail.Read'` with `config.AUTH_CONFIG.scopes.join(' ')` as default; add `offline_access` warning for `MS_SCOPES`
- [x] 2.3 `outlook-auth-server.js`: Replace inline `AUTH_CONFIG.scopes` array with `require('./config').AUTH_CONFIG.scopes`
- [x] 2.4 `auth/tools.js`: Migrate `handleCheckAuthStatus` — use `new TokenStorage().getValidAccessToken()` instead of `tokenManager.loadTokenCache()`

## Phase 3: REFACTOR & Verify

- [x] 3.1 Run `npm test` — all tests pass (existing + new)
- [x] 3.2 Verify backwards compatibility: existing token file format unchanged (no new fields in token JSON)
- [x] 3.3 Verify no circular dependencies: `auth/tools.js` does NOT require `auth/index.js` (uses `new TokenStorage()` directly)
