# Design: Fix Persistent OAuth Authentication

## Technical Approach

Eliminate the silent token-scope downgrade by making `config.js` the single source of truth for OAuth scopes. All scope consumers (`token-storage.js`, `outlook-auth-server.js`) import the unified scope list from `config.js` instead of defining inline lists. Token refresh sends the full scope set, and `check-auth-status` reads state from `TokenStorage` (which can refresh on demand) rather than the legacy `token-manager` (which reports expired tokens as "Not authenticated"). Maps to spec requirements: Scope Unification, Full-Scope Refresh, Auth Status Accuracy, offline_access Presence, Backwards Compatibility, One-Time Re-Auth.

## Architecture Decisions

### Decision: Unified scope list in config.js

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Union of all currently-used scopes in config.js | Adds Contacts.Read to tools that lacked it; one re-auth needed | ✅ Chosen |
| Keep per-module scope lists | No drift risk but defeats the fix | Rejected |

**Rationale**: The bug is scope fragmentation. A single union list (`offline_access Mail.Read Mail.ReadWrite Mail.Send User.Read Calendars.Read Calendars.ReadWrite Files.Read Files.ReadWrite Contacts.Read` — 10 scopes) ensures refresh requests exactly match initial auth. `offline_access` MUST be present or Microsoft returns no refresh token.

### Decision: MS_SCOPES override scope

| Option | Tradeoff | Decision |
|--------|----------|----------|
| MS_SCOPES overrides in token-storage.js only; warn if missing offline_access | Keeps power-user escape hatch; scoped to refresh path | ✅ Chosen |
| Remove MS_SCOPES entirely | Cleanest but breaks existing deployments | Rejected (proposal keeps it) |
| Apply MS_SCOPES to outlook-auth-server.js too | Spec forbids: auth server MUST use config.js directly | Rejected |

**Rationale**: Spec explicitly states `outlook-auth-server.js` MUST NOT use `MS_SCOPES`. The override stays as an explicit, validated escape hatch for `token-storage.js` only.

### Decision: check-auth-status TokenStorage instance

| Option | Tradeoff | Decision |
|--------|----------|----------|
| tools.js creates own `new TokenStorage()` | Avoids circular dep with auth/index.js; 2 instances share same file store | ✅ Chosen |
| Lazy-require singleton from auth/index.js | Same instance but hacky timing | Rejected |
| New shared singleton module | Cleanest but adds a file beyond proposal scope | Noted as future refactor |

**Rationale**: `auth/index.js` requires `./tools`; if `tools.js` required `./index` for the singleton it would create a load-time cycle. Two `TokenStorage` instances reading the same token file is safe — both load-on-demand from disk. `check-auth-status` only needs read+refresh, not cross-instance coordination.

### Decision: token-manager.js untouched

**Choice**: Leave `token-manager.js` as-is (no scopes, no refresh). It remains required for Flow tokens (`getFlowAccessToken`, `saveFlowTokens`) and test-mode token creation.
**Rationale**: Proposal marks it Out of Scope for removal; only `check-auth-status` migrates off it.

## Data Flow

    config.js (AUTH_CONFIG.scopes — 10 scopes, single source)
        │
        ├──→ outlook-auth-server.js  (scopes for /auth redirect + code exchange)
        │
        └──→ auth/token-storage.js    (scopes for refresh + code exchange)
                  │
                  └──→ getValidAccessToken()  ──→ refreshAccessToken() [full scopes in POST body]
                            │
                            └──→ auth/tools.js handleCheckAuthStatus() reads state here

    auth/token-manager.js  ──→ Flow tokens only (unchanged)

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `config.js` | Modify | Add `offline_access`, `Contacts.Read` to `AUTH_CONFIG.scopes` (10 total) |
| `auth/token-storage.js` | Modify | Default scopes from `config.AUTH_CONFIG.scopes`; keep `MS_SCOPES` override with `offline_access` warning |
| `outlook-auth-server.js` | Modify | Replace inline scopes array with `require('./config').AUTH_CONFIG.scopes`; keep own AUTH_CONFIG for clientId/tenant/authority |
| `auth/tools.js` | Modify | `handleCheckAuthStatus` uses `new TokenStorage().getValidAccessToken()` instead of `tokenManager.loadTokenCache()` |
| `auth/token-manager.js` | No change | Retained for Flow tokens; deprecation note optional |
| `test/auth/token-storage.test.js` | Modify | Add: default scopes from config.js, MS_SCOPES override, offline_access validation, refresh POST body full scopes |
| `test/auth/tools.test.js` | Create | New: check-auth-status with refreshable expired token → "Authenticated"; no tokens → "Not authenticated" |
| `auth/oauth-server.js` | No change | Out of scope (unused Express module) |

## Interfaces / Contracts

```js
// config.js — single source of truth
AUTH_CONFIG: {
  scopes: [
    'offline_access', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
    'User.Read', 'Calendars.Read', 'Calendars.ReadWrite',
    'Files.Read', 'Files.ReadWrite', 'Contacts.Read'
  ]
}

// auth/token-storage.js — default from config, override from env, validate
const config = require('../config');
const scopeString = process.env.MS_SCOPES || config.AUTH_CONFIG.scopes.join(' ');
if (process.env.MS_SCOPES && !scopeString.includes('offline_access')) {
  console.warn('MS_SCOPES override is missing offline_access — refresh tokens will not be issued.');
}
this.config.scopes = scopeString.split(' ');

// auth/tools.js — refresh-aware status
const TokenStorage = require('./token-storage');
async function handleCheckAuthStatus() {
  const tokenStorage = new TokenStorage();
  const token = await tokenStorage.getValidAccessToken();
  return { content: [{ type: "text", text: token ? "Authenticated and ready" : "Not authenticated" }] };
}
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | config.js scopes include `offline_access` + all 10 | Direct assertion on `config.AUTH_CONFIG.scopes` |
| Unit | token-storage default scopes come from config.js | Construct without MS_SCOPES; assert scopes match config |
| Unit | MS_SCOPES override + offline_access missing → warn | Set MS_SCOPES without offline_access; spy console.warn |
| Unit | refresh POST body includes full scope set | Mock https; parse querystring; assert scope equals config.join(' ') |
| Unit | check-auth-status expired+refreshable → "Authenticated" | Mock getValidAccessToken → token; assert text |
| Unit | check-auth-status no tokens → "Not authenticated" | Mock getValidAccessToken → null; assert text |
| Integration | Existing token file parses unchanged | Write legacy-format JSON; getTokens() returns it |

## Threat Matrix

N/A — no routing, shell, subprocess, VCS/PR automation, executable-file classification, or process-integration boundary.

## Migration / Rollout

No data migration. Existing token files remain readable (format unchanged). Users with downscoped tokens from the old refresh path MAY need one re-authentication to obtain a full-scope token (spec: One-Time Re-Auth). Document this in the `authenticate` tool output.

## Open Questions

- [ ] Should `auth/token-manager.js` get a deprecation JSDoc note directing new code to `token-storage`? (Non-blocking; optional in tasks.)