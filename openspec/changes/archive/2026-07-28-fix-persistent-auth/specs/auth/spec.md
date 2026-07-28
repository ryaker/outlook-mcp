# Auth — Persistent Authentication Specification

## Purpose

Define the behavior of the OAuth token lifecycle: scope configuration, token refresh, and auth status reporting. All auth modules MUST reference a single scope source in `config.js` to prevent silent scope downgrade on refresh.

## Requirements

### Requirement: Scope Unification

All auth modules that request scopes from Microsoft identity platform MUST reference `config.js` as the single source of truth for scope lists.

#### Scenario: All scope consumers reference config.js

- GIVEN `config.js` defines `AUTH_CONFIG.scopes` with the full scope set
- WHEN any auth module (`token-storage.js`, `outlook-auth-server.js`) constructs an OAuth request
- THEN it MUST use `config.AUTH_CONFIG.scopes` as the scope parameter
- AND it MUST NOT define its own inline scope list

#### Scenario: MS_SCOPES env var overrides config.js scopes

- GIVEN `process.env.MS_SCOPES` is set
- WHEN `token-storage.js` initializes its config
- THEN it MUST use `MS_SCOPES` instead of `config.AUTH_CONFIG.scopes`
- AND `outlook-auth-server.js` MUST NOT use `MS_SCOPES` (it uses `config.js` directly)

### Requirement: Full-Scope Token Refresh

`token-storage.refreshAccessToken()` MUST request the same scope set as the initial authorization code exchange.

#### Scenario: Refresh POST body includes full scopes

- GIVEN a stored refresh token from a previous auth with 10+ scopes
- WHEN `refreshAccessToken()` is called
- THEN the POST body to the token endpoint MUST include `scope` with all scopes from the configured list
- AND the `scope` parameter MUST NOT be a subset of the original auth scopes

#### Scenario: Refresh with downscoped env var override

- GIVEN `MS_SCOPES` is set to `"offline_access User.Read Mail.Read"`
- WHEN `refreshAccessToken()` is called
- THEN the POST body MUST use the `MS_SCOPES` value as the scope parameter
- AND the resulting token SHALL be limited to those scopes

### Requirement: Auth Status Accuracy

The `check-auth-status` tool MUST report authentication status based on `token-storage`'s token state, including its refresh capability.

#### Scenario: Token expired but refreshable reports authenticated

- GIVEN a stored token with `expires_at` in the past and a valid `refresh_token`
- WHEN `check-auth-status` is called
- THEN it MUST report "Authenticated" because `token-storage` can refresh the token on demand

#### Scenario: No tokens stored reports not authenticated

- GIVEN no token file exists at the configured path
- WHEN `check-auth-status` is called
- THEN it MUST report "Not authenticated"

### Requirement: offline_access Presence

The scope list in `config.js` MUST include `offline_access` to ensure Microsoft returns a refresh token during the authorization code exchange.

#### Scenario: offline_access in config.js scopes

- GIVEN `config.AUTH_CONFIG.scopes`
- WHEN the list is inspected
- THEN it MUST contain `"offline_access"`
- AND the initial auth URL and token exchange POST body MUST include `offline_access`

### Requirement: Backwards Compatibility

Existing code that reads tokens from the token file MUST continue to work after the scope unification.

#### Scenario: Token file format unchanged

- GIVEN a token file at `~/.outlook-mcp-tokens.json` created by the previous auth flow
- WHEN `token-storage.getTokens()` is called
- THEN it MUST parse and return the tokens successfully
- AND the token structure (`access_token`, `refresh_token`, `expires_at`, `expires_in`, `scope`) MUST be identical to the previous format

### Requirement: One-Time Re-Auth

After deploying the scope fix, users with existing tokens that were issued with a downscoped scope set MAY need to re-authenticate once to obtain a token with the full scope set.

#### Scenario: Existing downscoped token triggers re-auth prompt

- GIVEN a stored token whose `scope` field is a subset of `config.AUTH_CONFIG.scopes`
- WHEN `getValidAccessToken()` detects the token is expired and refresh returns a downscoped token
- THEN the system SHOULD surface a message indicating re-authentication may be needed
- AND the user MAY re-authenticate via the `authenticate` tool to obtain a full-scope token
