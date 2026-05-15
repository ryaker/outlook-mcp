# Design: Auth + multi-account token storage (v3)

**Slug:** `auth-multi-account`
**Date:** 2026-05-14
**Author:** inline design (Claude Opus 4.7)
**Status:** Ready for /decompose

---

## 1. Problem statement

The v3 rebuild needs an authentication layer from scratch. The v2 implementation we deleted had three structural limits that block what we need next:

1. **Single-account only.** Tokens were stored as a flat `{ access_token, refresh_token, expires_at }` object in one file. Re-authenticating as a different M365 account silently overwrote the previous tokens. There was no concept of "which account am I acting as right now."
2. **Single tenant per process.** `MS_TENANT_ID` was read once at startup. Accounts in different tenants (work + personal, or multiple work tenants) could not coexist.
3. **Hand-rolled OAuth.** ~500 LOC of `https.request` + manual `querystring` + manual expiry math + a custom concurrent-refresh dedup. Each reinvents a wheel Microsoft already ships in `@azure/msal-node`, and the v2 code had subtle correctness issues (e.g. token path resolved relative to `cwd` rather than a stable location, refresh-failure cleared in-memory tokens before persisting).

These cannot be patched into v2 without a near-total rewrite of `auth/`. That rewrite is this spec.

**Why it matters:** every other module — email, calendar, OneDrive, rules, Power Automate — depends on `getAuthenticatedGraphClient()`. Auth is the prerequisite for every subsequent design cycle in this rebuild. Multi-account in particular needs to land *before* any tool surface freezes, because every tool's input schema gains an optional `account` parameter and its output gains a mandatory `actedAs` field — and we don't want to retrofit either.

---

## 2. Proposed approach

Six components, in dependency order.

### 2.1 MSAL as the OAuth engine

Replace all hand-rolled OAuth with `@azure/msal-node`'s `PublicClientApplication`. MSAL handles:

- Authority URL construction per tenant (the `https://login.microsoftonline.com/{tenant}/...` plumbing).
- Token refresh with expiry buffer and concurrent-refresh deduplication.
- Multi-account in-memory cache keyed by `homeAccountId`.
- A serializable cache format (`tokenCache.serialize()` / `deserialize()`) that we can persist.

We use **public client** (no client secret), which is correct for delegated user-auth flows. The Azure app registration must be configured as a public client (Authentication blade → "Allow public client flows" → Yes). The v2 secret-based confidential client model is wrong for the device-code flow we're moving to.

### 2.2 Device-code flow as the only OAuth flow in MVP

`acquireTokenByDeviceCode()`. The user is given a short code and a `verificationUri` (`https://microsoft.com/devicelogin`); they complete the flow in a browser on any device. No callback URL needed — works in stdio, localhost HTTPS, headless containers identically.

The v2 authorization-code-with-localhost-callback pattern is rejected: device code is simpler, works the same way regardless of how the MCP is launched, and matches the headless-friendly profile we want.

Authorization-code-with-PKCE may be added later as a local-dev convenience but is out of scope for MVP.

### 2.3 TokenStore: keychain-when-available, file fallback

```ts
interface TokenStore {
  readonly name: string;            // "keychain" | "file" — for diagnostics
  load(): Promise<string | null>;   // returns serialized cache envelope, or null if empty
  save(serialized: string): Promise<void>;
  clear(): Promise<void>;
}
```

Two implementations:

- **`KeychainTokenStore`** — uses `@napi-rs/keyring` (prebuilt binaries, no `node-gyp`). One keychain entry, service `"outlook-mcp"`, account `"msal-cache"`, value = serialized cache envelope. Multi-account is handled inside the envelope, not by spreading across multiple keychain entries.
- **`FileTokenStore`** — writes the envelope to `$MCP_TOKEN_CACHE_PATH` (default `~/.outlook-mcp/cache.json`) with `0600` permissions. Parent dir auto-created at `0700`.

**Selection logic at startup** (`resolveTokenStore()`):

1. If `MCP_TOKEN_STORE` env var is set to `"keychain"` or `"file"`, use that explicitly. Throw if `"keychain"` is requested but unavailable.
2. Otherwise (default = `"auto"`): try `new KeychainTokenStore()` and a probe `load()` with a 2-second timeout. If it throws (no keyring daemon, no Secret Service on Linux, etc.) or times out, log to stderr that we fell back and instantiate `FileTokenStore`.
3. Selection happens once at process start; log the chosen backend to stderr.

This gives macOS/Windows local-dev users keychain security automatically, gives container/Linux deployments the file fallback automatically, and gives sysadmins an explicit override.

### 2.4 Account resolution — always implicit, default is visible

**Decision: every tool acts as a default account unless explicitly overridden.** This prioritizes ergonomics over per-call identity friction. The risk — losing track of which account is currently acting — is mitigated by **five visibility mechanisms** that make the default impossible to lose track of.

```ts
interface AccountSelector {
  resolve(requested: string | undefined): Promise<msal.AccountInfo>;
  // Always returns an account on success. Throws AccountResolutionError
  // with a user-friendly message + available UPNs on failure.
}
```

Resolution order:

1. If `requested` is provided: find an account whose `username` (UPN) matches case-insensitively. If none, throw with the list of available UPNs.
2. Else: use the **resolved default account**, computed as:
   - The **persisted runtime default** stored in the cache envelope (set via `auth_set_default_account`).
   - Else `MCP_DEFAULT_ACCOUNT` env var, validated against cached accounts.
   - Else the first-added account (by MSAL cache order — deterministic).
   - Else (zero accounts): throw `AuthenticationRequiredError` with "no accounts cached; run `auth_add_account` to begin."

#### Visibility mechanisms

1. **Every tool response echoes `actedAs` inside a `meta` envelope.** Every account-scoped tool's output includes a top-level `meta` object containing at minimum `{ actedAs: "<UPN>", actedAsTenantId: "<uuid>" }`. Even when `account` was implicit, the response makes it explicit. Auditable from the transcript alone. The `meta` envelope is also the extensibility point for future cross-cutting fields (request IDs, latency, rate-limit info).
2. **`auth_list_accounts` flags `isDefault: true`** on the resolved default account, with `defaultSource: "runtime" | "env" | "first-added"` so it's clear *why* it's the default.
3. **`auth_get_default_account` tool** — one-liner lookup returning `{ username, tenantId, defaultSource }`.
4. **`auth_set_default_account` tool** — change the runtime default without restarting; persists into the cache envelope.
5. **Server-startup log to stderr.** First line after transport connection: `Default account: shawn@outlook.com (from: env)` or `(from: first-added)` or `(none cached — run auth_add_account)`.
6. **Tool-description convention.** Every non-auth tool's MCP description (the text Claude reads when deciding to invoke a tool) must end with the standard sentence: *"Response includes `meta.actedAs` indicating which account was used."* This is a passive nudge — it surfaces the convention to Claude so it knows the field exists, without instruction-stuffing imperative language into every tool description. Per-module confirmation prompts for destructive operations (e.g., "About to send from X — proceed?") are out of scope here and called out in §8.

### 2.5 Five MCP tools + one internal helper

**Public tools (registered on the MCP server):**

| Tool | Purpose |
|---|---|
| `auth_add_account` | Initiates device-code flow. Returns `{ userCode, verificationUri, expiresInSeconds, message }` immediately. MSAL polls in the background; the account appears in `auth_list_accounts` once the user completes the browser flow. |
| `auth_list_accounts` | Returns `{ accounts: [{ username, tenantId, homeAccountId, isDefault, defaultSource? }], pendingAdditions: [{ userCode, verificationUri, expiresAt }] }`. |
| `auth_remove_account` | Takes `{ account: string (UPN) }`. Removes from MSAL cache, persists. Idempotent — `removed: false` if not present. If the removed account was the persisted runtime default, the runtime default is cleared (subsequent calls fall back to env → first-added). |
| `auth_get_default_account` | Returns the currently-resolved default: `{ username, tenantId, defaultSource: "runtime" \| "env" \| "first-added" } \| null` (null if zero accounts cached). |
| `auth_set_default_account` | Takes `{ account: string (UPN) }`. Sets the persisted runtime default. Throws if UPN doesn't match a cached account. Returns the new default's `{ username, defaultSource: "runtime" }`. |

**Internal helper (used by all future modules):**

```ts
export async function getAuthenticatedGraphClient(
  account?: string,
): Promise<{ client: Client; actedAs: string }>;
```

Resolves the account, acquires a token silently via MSAL (`acquireTokenSilent`), and returns a `@microsoft/microsoft-graph-client` `Client` configured with that bearer token, plus the UPN that was used (so the calling tool can include it in its `actedAs` response field). On silent-acquisition failure (e.g., refresh token expired), throws `AuthenticationRequiredError` with the account UPN so the caller can surface "please re-run `auth_add_account` for this account."

### 2.6 No v2 token migration

The v2 token file at `~/outlook-mcp-tokens.json.backup-2026-05-14` is a flat single-account blob with no `homeAccountId`, no tenant info, no MSAL metadata. Reconstructing an MSAL cache entry from it is fragile — we'd have to decode the access token to recover account info, and the access token has likely expired anyway. The migration cost is one `auth_add_account` call by the user. We skip it.

---

## 3. Rejected alternatives

### 3.1 Keep hand-rolled OAuth, just add a per-account map

Considered: smaller diff from v2, no new dependency. Rejected because:
- The v2 OAuth code has known correctness issues (concurrent-refresh, token-path resolution, refresh-failure ordering).
- We'd still need to invent multi-tenant URL routing, per-account scope handling, and a concurrent-refresh-per-account dedup. MSAL ships all of this.
- The v3 rebuild's whole point is to stop reinventing wheels Microsoft maintains.

### 3.2 Authorization-code flow with local callback server

Considered: most familiar OAuth shape, what v2 used. Rejected because:
- Requires a callback URL the browser can reach. Device code works identically across stdio, localhost HTTP, and any hosting topology without infrastructure dependencies.
- We can add PKCE later if the local-dev UX of "browser auto-opens" is missed.

### 3.3 Keychain-only storage

Considered: most-secure local storage. Rejected because:
- Headless Linux containers (Alpine, distroless) have no D-Bus/Secret Service. `@napi-rs/keyring` throws on first call. This breaks portability.
- "Auto-detect with file fallback" gets us the security benefit on Mac/Windows local dev *and* a working fallback for any other context, with one extra startup probe.

### 3.4 Per-account file (`~/.outlook-mcp/<homeAccountId>.json`)

Considered: avoids the single-file-corruption-risk-on-concurrent-write concern. Rejected because:
- MSAL doesn't support per-account serialization — its cache is monolithic. Splitting it would mean re-implementing MSAL's cache itself.
- We're a single-process server; no concurrent writes to worry about.

### 3.5 Always require `account` explicitly (Google Workspace MCP's pattern)

Considered: forces explicit, leaves no ambiguity. Rejected because:
- Single-account is the common case for individual users; typing `account: "shawn@outlook.com"` on every call adds friction without payoff in that case.
- The five visibility mechanisms (especially the `actedAs` echo in every tool response, and the startup-stderr log) provide the audit trail that the always-explicit pattern is meant to give, without the per-call friction.

### 3.6 Implicit when one, explicit when many

Considered as a middle ground (and the original draft of this spec). Rejected because:
- The hard failure at the moment of ambiguity is worse UX than picking a defensible default and surfacing it loudly via the visibility mechanisms.
- With the persisted runtime default + the env override, the user has explicit control over the default without needing the "force explicit" stick.

### 3.7 Auto-encrypt the file-store at rest

Considered: closes the "lost laptop without FileVault" gap. Rejected for MVP because:
- Need a key. Storing the key on disk next to the encrypted file is theater; pulling it from the OS keychain reproduces the keychain dependency we just fell back from.
- The keychain backend already provides at-rest encryption on the platforms where it matters most.
- Can be added via the `TokenStore` interface later (e.g., `EncryptedFileTokenStore` wrapping `FileTokenStore` with a key from `MCP_CACHE_KEY` env).

### 3.8 Network-backed TokenStore (Redis / Key Vault) for "Cowork hosting"

Considered: original spec draft assumed Cowork hosted the MCP in an ephemeral container, motivating a remote store. Rejected because **the premise was wrong**: research confirmed Cowork is a local desktop app, not a cloud hosting plane (see §6.1). For this deployment model, the local filesystem persists like any user file. A network-backed store has no purpose in the local-only deployment model and is correctly an *out-of-scope* item for any future public-hosted variant (§8).

---

## 4. Affected files

| Path | Change | Reason |
|---|---|---|
| `package.json` | modify | Add deps: `@azure/msal-node`, `@microsoft/microsoft-graph-client`, `@napi-rs/keyring`. |
| `.env.example` | modify | Document `MS_CLIENT_ID`, `MS_TENANT_ID` (default `common`), `MCP_TOKEN_STORE` (`auto`/`keychain`/`file`), `MCP_TOKEN_CACHE_PATH`, `MCP_DEFAULT_ACCOUNT`. Remove `MS_CLIENT_SECRET` (public client). |
| `src/config.ts` | create | Centralized config — reads env, validates with Zod, exposes typed config object. |
| `src/auth/errors.ts` | create | `AuthenticationRequiredError`, `AccountResolutionError`, `TokenStoreUnavailableError`. |
| `src/auth/cache-envelope.ts` | create | Schema-versioned wrapper `{ schemaVersion: 1, defaultAccount: string \| null, msalCache: string }`. Encode/decode + future-migration hook. |
| `src/auth/token-store.ts` | create | `TokenStore` interface + `resolveTokenStore()` factory. |
| `src/auth/stores/keychain-store.ts` | create | `KeychainTokenStore` using `@napi-rs/keyring`. |
| `src/auth/stores/file-store.ts` | create | `FileTokenStore` with `0600`/`0700` permissions and parent-dir creation. |
| `src/auth/msal-client.ts` | create | Builds and memoizes `PublicClientApplication` with the configured `TokenStore` wired into `cachePlugin`. Exposes `getMsal(): PublicClientApplication` and `getRuntimeDefaultAccount()` / `setRuntimeDefaultAccount(upn)`. |
| `src/auth/account-selector.ts` | create | `resolveAccount(requested?)`. Implements the resolution order in §2.4. |
| `src/auth/device-code-flow.ts` | create | Orchestrates `acquireTokenByDeviceCode`, holds the in-flight promise map keyed by user code, surfaces pending state to `auth_list_accounts`. |
| `src/auth/tools.ts` | create | Five MCP tools with Zod-validated inputs. Exports `authTools: McpTool[]`. |
| `src/auth/startup-log.ts` | create | Emits the default-account stderr line on server start. |
| `src/auth/index.ts` | create | Barrel: re-exports `authTools`, `getAuthenticatedGraphClient`, error classes. |
| `src/graph/client.ts` | create | `getAuthenticatedGraphClient(account?)` — wraps `@microsoft/microsoft-graph-client` with the MSAL-acquired token, returns `{ client, actedAs }`. |
| `src/server.ts` | modify | Drop the `ping` stub. Register `authTools`. Invoke `logDefaultAccount()` after transport connects. |
| `src/index.ts` | unchanged | Transport dispatch already in place. |
| `src/auth/__tests__/account-selector.test.ts` | create | Unit tests for resolution rules. |
| `src/auth/__tests__/file-store.test.ts` | create | Unit tests for file store: round-trip, permissions, missing-dir creation, corruption handling. |
| `src/auth/__tests__/keychain-store.test.ts` | create | Unit tests with a mocked keyring. |
| `src/auth/__tests__/token-store-resolver.test.ts` | create | Tests for the auto-detect / explicit-env logic. |
| `src/auth/__tests__/cache-envelope.test.ts` | create | Round-trip, version detection, malformed-input handling. |
| `src/auth/__tests__/tools.test.ts` | create | Tests for the five MCP tools with mocked MSAL. |

---

## 5. Data & interface changes

### 5.1 Environment variables

| Var | Required | Default | Purpose |
|---|---|---|---|
| `MS_CLIENT_ID` | yes | — | Azure app registration client ID. Public client (no secret). |
| `MS_TENANT_ID` | no | `common` | Authority tenant. Per-account tenant is detected from the token. |
| `MCP_TOKEN_STORE` | no | `auto` | `auto` \| `keychain` \| `file`. |
| `MCP_TOKEN_CACHE_PATH` | no | `~/.outlook-mcp/cache.json` | File-store location (only used if file store selected). |
| `MCP_DEFAULT_ACCOUNT` | no | — | UPN used as default when no persisted runtime default exists. |
| `MCP_TRANSPORT` | no | `stdio` | (Already in scaffold; unchanged.) |

`MS_CLIENT_SECRET` is removed — public-client device-code flow doesn't use it.

### 5.2 Cache file shape (schema-versioned wrapper)

```json
{
  "schemaVersion": 1,
  "defaultAccount": "shawn@outlook.com",
  "msalCache": "<stringified MSAL cache JSON>"
}
```

`defaultAccount` is the persisted runtime default set via `auth_set_default_account`. Null/absent means "no runtime default set; fall through to env → first-added."

`msalCache` is the raw string form of MSAL's serialized cache. We don't parse or modify it — MSAL owns its format.

`schemaVersion` is an integer. If we ever need to migrate (MSAL major bump, or our own envelope shape change), `cache-envelope.ts` switches on this version. Versions <1 (legacy raw MSAL) are detected by absence of the wrapper object and surfaced as a user-facing "please re-auth" message rather than silently migrated.

### 5.3 Public TypeScript surface (`src/auth/index.ts`)

```ts
export { authTools } from "./tools.js";
export { getAuthenticatedGraphClient } from "../graph/client.js";
export {
  AuthenticationRequiredError,
  AccountResolutionError,
  TokenStoreUnavailableError,
} from "./errors.js";
```

### 5.4 Tool input/output schemas (Zod-defined, JSON-Schema-rendered for MCP)

**`auth_add_account`**

```ts
const Input = z.object({});

const Output = z.object({
  userCode: z.string(),
  verificationUri: z.string().url(),
  expiresInSeconds: z.number().int(),
  message: z.string(),
});
```

**`auth_list_accounts`**

```ts
const Input = z.object({});

const Output = z.object({
  accounts: z.array(z.object({
    username: z.string(),
    tenantId: z.string().uuid(),
    homeAccountId: z.string(),
    isDefault: z.boolean(),
    defaultSource: z.enum(["runtime", "env", "first-added"]).optional(), // present when isDefault: true
  })),
  pendingAdditions: z.array(z.object({
    userCode: z.string(),
    verificationUri: z.string().url(),
    expiresAt: z.string().datetime(),
  })),
});
```

**`auth_remove_account`**

```ts
const Input = z.object({
  account: z.string().min(1),
});

const Output = z.object({
  removed: z.boolean(),
  remainingAccounts: z.array(z.string()),
  defaultClearedAsResult: z.boolean(),  // true if we just removed the persisted runtime default
});
```

**`auth_get_default_account`**

```ts
const Input = z.object({});

const Output = z.object({
  account: z.object({
    username: z.string(),
    tenantId: z.string().uuid(),
    defaultSource: z.enum(["runtime", "env", "first-added"]),
  }).nullable(),  // null when zero accounts cached
});
```

**`auth_set_default_account`**

```ts
const Input = z.object({
  account: z.string().min(1),
});

const Output = z.object({
  username: z.string(),
  defaultSource: z.literal("runtime"),
});
```

### 5.5 Cross-cutting convention for all future tools

Every non-auth tool, in every future module, **must**:

1. Accept an optional top-level `account: string` field in its Zod input schema.
2. Resolve the account via `getAuthenticatedGraphClient(input.account)` rather than reading from a singleton.
3. Include a top-level `meta` object in its output schema with at minimum `{ actedAs: string, actedAsTenantId: string }`. The `meta` envelope is the extensibility point — future cross-cutting fields (request IDs, latency, rate-limit info) live here. Tool-specific payload fields live alongside `meta` at the top level, not inside it.
4. On errors, when an account was successfully resolved *before* the failure occurred, include `actedAs` (and ideally `actedAsTenantId`) in the error's structured `data` field so the caller can tell *which* account attempted the operation. Errors thrown before account resolution (e.g., `AuthenticationRequiredError` from zero-account state) may omit this.
5. End the tool's MCP description with the standard sentence: *"Response includes `meta.actedAs` indicating which account was used."*

This convention is **locked here**, in the auth design, so every subsequent module's design references it rather than re-deriving. Retrofitting `meta.actedAs` later would touch every tool.

### 5.6 Initial scope set

```
offline_access User.Read
Mail.ReadWrite Mail.Send
Calendars.ReadWrite
Contacts.Read
Files.ReadWrite Files.ReadWrite.All
MailboxSettings.ReadWrite
```

Covers the v3 module plan (email, calendar, folders, rules, OneDrive). Redundant `.Read` variants trimmed (`Mail.ReadWrite` supersedes `Mail.Read`, etc.). Power Automate's scopes (`https://service.flow.microsoft.com//.default`) are a separate namespace and require a separate consent — deferred to the Power Automate module's design.

Adding a scope later requires re-consent (the user goes through `auth_add_account` again); Microsoft's incremental-consent UX shows only the *new* scopes on that second run.

---

## 6. Risks & open questions

Numbered for traceability in the QA report.

### 6.1 [RESOLVED] Cowork deployment shape

**Original concern:** if Cowork hosts the MCP in an ephemeral container with no persistent volume, the `FileTokenStore` loses tokens on every restart.

**Resolution:** the premise was wrong. Research (2026-05-14, primary sources at `platform.claude.com/docs/en/agents-and-tools/mcp-connector`, `support.claude.com/articles/13345190`, `support.claude.com/articles/13854387`, `support.claude.com/articles/11175166`) confirmed:
- Cowork is a local desktop app — runs on the user's machine, not on Anthropic infrastructure.
- Scheduled tasks "only run while your computer is awake and the Claude Desktop app is open."
- Remote MCP servers are user-hosted at user-controlled HTTPS endpoints; Anthropic does not provide hosting.

For our deployment (local stdio spawned by Cowork, or optional localhost Streamable HTTP), the cache file is on the user's home directory and persists like any other file. **No network-backed TokenStore is needed.**

The laptop-on / app-open constraint is documented as product reality in §6.9, not as a design risk.

### 6.2 [RESOLVED] HTTP transport inbound auth

**Original concern:** if the server is reachable over HTTP, anyone who can hit it can use cached accounts.

**Resolution:** for local-only deployment (stdio or localhost HTTPS), the server is not exposed to the network. Inbound auth is moot. The MCP spec (2025-11-25, `modelcontextprotocol.io/specification/2025-11-25/basic/authorization`) mandates OAuth 2.1 resource-server semantics for any *publicly-hosted* MCP — that is a known path and is correctly out of scope for this slug (see §8).

### 6.3 [RISK] MSAL cache format coupling

*Likelihood:* low.
*Impact:* low — recoverable by detecting old format and prompting re-auth.

Storing MSAL's serialized cache (inside our envelope) couples us to MSAL's schema. If MSAL changes the format between majors, users on the old format need to re-auth.

*Mitigation:* the `schemaVersion: 1` wrapper makes detection trivial. If we ever bump to `schemaVersion: 2` to accommodate MSAL changes, `cache-envelope.ts` returns a typed error with a "please re-run `auth_add_account` for any cached accounts" message. No silent migration of credentials.

### 6.4 [RISK] Device-code flow timeout

*Likelihood:* medium.
*Impact:* low — visible failure with clear message.

MSAL's device-code polling has a default timeout (~15 min). If the user doesn't complete in time, the in-flight promise rejects.

*Mitigation:* wrap the MSAL promise with explicit timeout cleanup; on rejection, remove from `pendingAdditions` and log to stderr. `auth_list_accounts` users see the pending entry disappear with no further explanation, which is fine — they can simply re-run `auth_add_account`.

### 6.5 [RISK] Concurrent `auth_add_account` calls

*Likelihood:* low.
*Impact:* low.

If a user calls `auth_add_account` twice without completing the first, we have two device codes outstanding. Both will be tracked in `pendingAdditions`; whichever the user completes first wins, the other times out.

*Mitigation:* acceptable as designed; no special handling needed.

### 6.6 [RISK] Keychain probe blocking on macOS first run

*Likelihood:* low.
*Impact:* medium — first-run keychain prompt may interrupt non-interactive use.

On macOS, the first `KeychainTokenStore.load()` may trigger an "Allow outlook-mcp to access keychain?" prompt. If the user is running headless via SSH or a launchd agent, the prompt blocks.

*Mitigation:* the probe in `resolveTokenStore()` uses a 2s timeout; if it doesn't resolve, fall back to file. Document the keychain prompt in the README so first-run users aren't surprised.

### 6.7 [RISK] Token-cache write race in async tool calls

*Likelihood:* very low (Node is single-threaded; MSAL serializes cache mutations).
*Impact:* low.

Multiple concurrent tool calls could in principle trigger overlapping `save()` calls.

*Mitigation:* MSAL's `cachePlugin` is invoked via `beforeCacheAccess`/`afterCacheAccess` hooks that serialize within MSAL's own locking. We trust that. File writes are atomic via `fs.writeFile` on POSIX.

### 6.8 [RISK] `defaultAccount` referencing a removed account

*Likelihood:* medium.
*Impact:* low.

If the user manually edits the cache file or some other corruption happens, the persisted `defaultAccount` UPN could refer to an account no longer in the MSAL cache.

*Mitigation:* `account-selector.ts` always validates the persisted default against the current cached set before using it. If invalid: fall through to env → first-added, and log a stderr line. `auth_remove_account` already clears the runtime default if removing the matching account (see §2.5).

### 6.9 [PRODUCT REALITY] Laptop-on, app-open

Not a design risk per se, but worth documenting: Cowork scheduled tasks require the user's computer to be awake and Claude Desktop to be open at the scheduled time. The user has explicitly accepted this constraint (2026-05-14, design check-in). If a future need surfaces for truly-headless scheduling ("run at 3am while I'm asleep"), the right answer is **not** to change this MCP — it's to host it publicly and trigger from a different surface (cloud cron, Claude Code Dispatch, etc.). That's the future slug outlined in §8.

---

## 7. Test strategy

Coverage targets, by file. All tests use Vitest. No real network or real keychain access in CI.

### 7.1 Unit tests

**`account-selector.test.ts`** — 9 cases:
1. Single cached account, no requested, no persisted default, no env → returns that account, `defaultSource: "first-added"`.
2. Multiple cached, no requested, no persisted, no env → returns first-added, `defaultSource: "first-added"`.
3. Multiple cached, persisted default matches → returns persisted, `defaultSource: "runtime"`.
4. Multiple cached, persisted default doesn't match any (stale) → falls back to env / first-added, logs warning.
5. Multiple cached, env matches → returns env account, `defaultSource: "env"`.
6. Multiple cached, env doesn't match any → falls back to first-added with warning.
7. Multiple cached, requested matches case-insensitively → returns it, `defaultSource: "explicit-request"` (not persisted).
8. Multiple cached, requested doesn't match → throws with UPN list.
9. Zero cached → throws `AuthenticationRequiredError` "no accounts; run `auth_add_account`."

**`file-store.test.ts`** — 6 cases:
1. `save()` then `load()` round-trips identically.
2. `load()` returns `null` when file doesn't exist.
3. `save()` creates parent dir with `0700` if missing.
4. `save()` writes file with `0600` permissions (verify via `stat`).
5. `clear()` removes file; subsequent `load()` returns `null`.
6. `load()` on corrupted JSON throws a typed error.

**`keychain-store.test.ts`** — 5 cases (mocked `@napi-rs/keyring`):
1. `save()` then `load()` round-trips.
2. `load()` returns `null` when entry doesn't exist.
3. Mocked keyring throwing on `getPassword` → wraps in `TokenStoreUnavailableError`.
4. `clear()` deletes the entry.
5. Service/account names are exactly `"outlook-mcp"` / `"msal-cache"`.

**`token-store-resolver.test.ts`** — 5 cases:
1. `MCP_TOKEN_STORE=keychain` + keychain available → returns `KeychainTokenStore`.
2. `MCP_TOKEN_STORE=keychain` + keychain throws → re-throws (no silent fallback when explicit).
3. `MCP_TOKEN_STORE=file` → returns `FileTokenStore` even if keychain available.
4. `MCP_TOKEN_STORE=auto` + keychain probe succeeds → returns `KeychainTokenStore`.
5. `MCP_TOKEN_STORE=auto` + keychain probe times out → returns `FileTokenStore`, logs fallback.

**`cache-envelope.test.ts`** — 5 cases:
1. Round-trip `{schemaVersion: 1, defaultAccount, msalCache}` encode/decode preserves all fields.
2. Decode of raw MSAL JSON (no wrapper) → returns typed error suggesting re-auth, not silent migration.
3. Decode of `schemaVersion: 999` (unknown future) → returns typed error.
4. Encode with `defaultAccount: null` → field present as `null` in output, not omitted.
5. Decode of malformed JSON → typed error, not raw `SyntaxError`.

### 7.2 Integration tests (mocked MSAL)

**`tools.test.ts`** — 13 cases:
1. `auth_add_account` returns a `userCode` + `verificationUri` from the mocked MSAL device-code callback.
2. `auth_add_account` registers the pending addition; `auth_list_accounts` shows it.
3. When mocked MSAL device-code resolves, `auth_list_accounts` shows the new account and removes from pending.
4. When mocked MSAL device-code rejects (timeout), `auth_list_accounts` no longer shows the pending addition.
5. `auth_list_accounts` with one account marks it `isDefault: true, defaultSource: "first-added"`.
6. `auth_list_accounts` with multiple, `MCP_DEFAULT_ACCOUNT` matching one → that one `isDefault: true, defaultSource: "env"`.
7. `auth_list_accounts` after `auth_set_default_account` → persisted account `isDefault: true, defaultSource: "runtime"`.
8. `auth_remove_account` for an existing UPN → returns `removed: true`, account gone.
9. `auth_remove_account` for a non-existent UPN → returns `removed: false` without error.
10. `auth_remove_account` for the persisted runtime default → returns `defaultClearedAsResult: true`, subsequent `auth_get_default_account` reflects the fall-through.
11. `auth_get_default_account` with zero accounts → returns `{ account: null }`.
12. `auth_set_default_account` with non-existent UPN → throws with available UPNs in message.
13. `auth_set_default_account` then `auth_get_default_account` round-trip → matches.

### 7.3 Manual / acceptance

Performed by Shawn before merging the implementation:
1. Fresh install: `npm install && npm run build && npm run start:stdio`, send `auth_add_account` via MCP Inspector, complete in browser, verify token cache exists at expected location.
2. Run `auth_list_accounts` — shows the account with correct UPN, tenant, and `isDefault: true, defaultSource: "first-added"`.
3. Restart server — `auth_list_accounts` still shows the account (persistence works); stderr shows `Default account: ... (from: first-added)`.
4. Run `auth_add_account` for a second account, complete it; `auth_list_accounts` shows both, first one still default.
5. Run `auth_set_default_account` with the second's UPN; restart server; verify stderr shows `(from: runtime)` and second account is now default.
6. Run `auth_remove_account` on the default; verify `defaultClearedAsResult: true` and stderr on next restart shows fall-through.
7. Set `MCP_TOKEN_STORE=file` explicitly; verify a fresh `~/.outlook-mcp/cache.json` is created with `0600` (no keychain interaction).
8. On macOS with default `MCP_TOKEN_STORE=auto`: verify keychain prompt on first run, then no prompt on subsequent runs.

---

## 8. Out of scope (explicit non-goals)

- **Power Automate / Flow API auth** — separate scope namespace, cut as its own slug later.
- **Authorization-code-with-PKCE flow** — defer; device code covers MVP.
- **Application permissions / client credentials flow** — for service-principal scenarios; not relevant for MVP.
- **Token encryption at rest beyond keychain's coverage** — `EncryptedFileTokenStore` is a future v-next.
- **Cross-machine token roaming** — out of scope; user re-auths per device.
- **A `cancel_pending_addition` tool** — add only if it becomes painful.
- **The Graph client itself beyond a minimal `getAuthenticatedGraphClient()` factory** — pagination, retry, batching, ODATA helpers are deferred to the first module that needs them (likely email).
- **Public/multi-user hosting of `outlook-mcp`** — would require (a) a network-backed `TokenStore` implementation (the interface already accommodates this), (b) OAuth 2.1 resource-server semantics for inbound auth per the MCP spec, (c) per-request bearer-token-to-account-identity mapping, (d) different deployment topology. Future slug name: `outlook-mcp-public-hosting`. Not blocking the rebuild.
- **Truly-headless scheduled execution** (laptop closed) — Cowork doesn't support this; if needed, the user hosts `outlook-mcp` publicly per the above and triggers from a cloud cron / Claude Code Dispatch surface.
- **Per-module confirmation prompts for destructive operations.** High-stakes actions like `send_email`, `delete_calendar_event`, `delete_file` may benefit from a "About to do X as account Y — proceed?" prompt that surfaces the acting account in human-readable form before execution. This is a per-module design choice (belongs in the design for each destructive tool), not an auth-layer responsibility. The auth layer provides the data (`meta.actedAs`); whether and how each module prompts is decided when that module is designed.

---

## 9. Hand-off checklist

After approval, `/decompose auth-multi-account` should produce a TASKS.md with roughly this granularity:

1. Add deps (`@azure/msal-node`, `@microsoft/microsoft-graph-client`, `@napi-rs/keyring`) + update `.env.example`.
2. Implement `src/config.ts` (Zod-validated env config).
3. Implement `src/auth/errors.ts`.
4. Implement `src/auth/cache-envelope.ts` + tests.
5. Implement `FileTokenStore` + tests.
6. Implement `KeychainTokenStore` + tests.
7. Implement `resolveTokenStore()` + tests.
8. Implement `msal-client.ts` (PublicClientApplication factory with cache plugin and runtime-default getter/setter).
9. Implement `account-selector.ts` + tests.
10. Implement `device-code-flow.ts` + pending-additions map.
11. Implement `getAuthenticatedGraphClient()` in `src/graph/client.ts`.
12. Implement the five MCP tools in `src/auth/tools.ts` + tests.
13. Implement `startup-log.ts`.
14. Wire `authTools` + startup log into `src/server.ts`; drop `ping`.
15. README updates. Must include:
    - **Azure app registration walkthrough** — step-by-step:
        1. portal.azure.com → App registrations → New registration.
        2. **Supported account types: "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts."** This is the most common first-time tripwire — without the "+ personal Microsoft accounts" half, `@outlook.com` / `@hotmail.com` / `@live.com` accounts get rejected at consent with a cryptic error. Single users with only one work account can use single-tenant instead, but multi-account users almost certainly want this.
        3. **Authentication blade → "Allow public client flows" → Yes.** Without this, MSAL's device-code flow is rejected.
        4. **No client secret.** Public client — don't generate one. `MS_CLIENT_SECRET` is intentionally not in §5.1.
        5. Copy the Application (client) ID into `MS_CLIENT_ID`. Keep `MS_TENANT_ID=common` (default) unless single-tenant.
        6. API permissions — *technically optional* (device-code first-consent will offer scopes dynamically), but pre-adding the scopes from §5.6 produces a cleaner first-time consent UX.
       Without these settings, first-time users get past `npm install` and hit a wall at `auth_add_account`.
    - Auth section (the five tools, when to use each, the meta.actedAs convention).
    - Env table (all variables in §5.1).
    - Multi-account UX walkthrough (add → list → set-default → use → remove).
    - Troubleshooting (keychain prompt on macOS, token-cache location, common MSAL errors, force-fallback via `MCP_TOKEN_STORE=file`).
    - Laptop-on / app-open caveat for Cowork scheduled tasks (§6.9).
16. Manual acceptance run-through (per §7.3).
