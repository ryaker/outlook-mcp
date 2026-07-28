# Proposal: Fix Persistent OAuth Authentication

## Intent

Users re-authenticate every few hours because token refresh requests a downscoped token. `token-storage.js` uses 3 scopes (`offline_access User.Read Mail.Read`) on refresh, while the original token had 10. Microsoft returns a token with *requested* scopes, silently dropping mail send, calendar, contacts, files, and mail read-write.

## Scope

### In Scope
- Unify all 4 scope lists to single source of truth in `config.js`
- Fix `token-storage.js` to reference `config.js` scopes
- Fix `outlook-auth-server.js` to reference `config.js` scopes
- Migrate `check-auth-status` to use `token-storage` (not `token-manager`)
- Add `offline_access` and `Contacts.Read` to `config.js`
- Tests for scope consistency and correct refresh scopes

### Out of Scope
- Rewriting auth server or OAuth flow
- Removing `token-manager.js` (needed for Flow tokens)
- Changing `MS_SCOPES` env var override
- Fixing `auth/oauth-server.js` (unused Express module)

## Capabilities

### New Capabilities
- `auth-scope-config`: Centralized scopes in `config.js` consumed by all auth modules
- `auth-token-refresh`: Token refresh that requests full scope set, preventing silent downgrade

### Modified Capabilities
None — no existing auth specs in `openspec/specs/`.

## Approach

1. **`config.js`**: Add `offline_access`, `Contacts.Read` to `AUTH_CONFIG.scopes`
2. **`auth/token-storage.js`**: Import scopes from `config.js`; keep `MS_SCOPES` as override
3. **`outlook-auth-server.js`**: Import scopes from `config.js`
4. **`auth/tools.js`**: `handleCheckAuthStatus` uses `token-storage` singleton for accurate status
5. **Tests**: Scope consistency + refresh POST body verification

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `config.js` | Modified | Add `offline_access`, `Contacts.Read` |
| `auth/token-storage.js` | Modified | Import scopes from `config.js` |
| `outlook-auth-server.js` | Modified | Import scopes from `config.js` |
| `auth/tools.js` | Modified | Use `token-storage` for `check-auth-status` |
| `test/auth/token-storage.test.js` | Modified | Update for new scope source |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Existing tokens may need re-auth after fix | High | Document one re-auth may be needed |
| `MS_SCOPES` env var could still cause drift | Low | Keep as explicit override, document precedence |
| Flow token path accidentally broken | Low | No changes to `token-manager.js` |

## Rollback Plan

Revert scope changes in `config.js`, `token-storage.js`, `outlook-auth-server.js`. Revert `tools.js` to `token-manager`. Run `npm test`.

## Dependencies

None.

## Success Criteria

- [ ] All 4 scope consumers reference `config.js` as single source of truth
- [ ] `token-storage.refreshAccessToken()` sends full scope set in POST body
- [ ] `check-auth-status` reports accurate status when token-manager reports expired
- [ ] `npm test` passes
