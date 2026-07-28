# Design: Fix `move-emails` to support subfolder targets

## Technical Approach

Enhance `getFolderIdByName(accessToken, folderName)` in `email/folder-utils.js` to split `folderName` on `/`, traverse the folder hierarchy level-by-level, and return the leaf folder ID. Single-segment inputs (no `/`) pass through the existing top-level lookup unchanged — zero blast radius for flat names. Each subsequent segment queries `me/mailFolders/{parentId}/childFolders` with the same exact-match + case-insensitive-fallback pattern already used for top-level. A small internal helper isolates the per-segment resolution so the exact-match/fallback logic is not duplicated.

Maps to proposal Approach 1 and spec requirements: Path Resolution, Backwards Compatibility, Case-Insensitive Fallback, Error Handling, Empty Segment Filtering, Literal Slash Limitation.

## Architecture Decisions

| Decision | Option | Tradeoff | Choice | Rationale |
|----------|--------|----------|--------|-----------|
| Path splitting | (a) split+traverse in `getFolderIdByName` (b) new `resolveFolderPathHierarchy()` | (a) one function, all 4 call sites benefit transparently; (b) cleaner separation but requires call-site changes | (a) inline in `getFolderIdByName` | Proposal demands all call sites benefit with no schema changes; signature stays `(accessToken, folderName) → string\|null` |
| Per-segment resolver | (a) inline exact+fallback at each level (b) extract `resolveSegmentInParent(accessToken, parentId, segment)` | (a) duplicates ~20 lines 3x; (b) one helper, testable in isolation | (b) extract helper | DRY; matches existing exact-then-fallback pattern; easier unit testing per level |
| Empty segment handling | (a) `split('/').filter(s => s.trim())` (b) regex split | (a) simple, readable; (b) overkill | (a) `split('/').map(trim).filter(Boolean)` | Handles `//`, leading `/`, trailing `/` per spec Empty Segment Filtering |
| Child query `$top` | (a) omit (b) `$top=100` | (a) Graph default ~10; (b) matches existing fallback pattern | (b) `$top=100` on fallback call | Consistency with existing top-level fallback; proposal risk note |
| Mock-data scope | (a) only unit-test mocks (b) also extend `mock-data.js` | (a) tests pass; (b) `npm run test-mode` server also works with paths | (b) both | Proposal affected-areas lists `utils/mock-data.js`; keeps test-mode server functional |

## Data Flow

```
caller (move.js / create.js / rules/create.js)
   │
   ▼
getFolderIdByName(accessToken, "Tramite/REQ-104951")
   │
   ├─ split('/') → ["Tramite", "REQ-104951"]
   ├─ segment[0] "Tramite" ── resolveSegmentInParent(null, "Tramite")
   │        ├─ GET me/mailFolders?$filter=displayName eq 'Tramite'   (exact)
   │        └─ [fallback] GET me/mailFolders?$top=100                 (case-insensitive)
   │        ⇒ parentId = "tramite-id"
   ├─ segment[1] "REQ-104951" ── resolveSegmentInParent("tramite-id", "REQ-104951")
   │        ├─ GET me/mailFolders/{parentId}/childFolders?$filter=displayName eq 'REQ-104951'
   │        └─ [fallback] GET me/mailFolders/{parentId}/childFolders?$top=100
   │        ⇒ leafId = "req-104951-id"
   └─ return leafId
```

## Interfaces / Contracts

**Public — unchanged signature:**
```js
// email/folder-utils.js
async function getFolderIdByName(accessToken, folderName)
// Returns: Promise<string|null>  — leaf folder ID, or null if any segment unresolved
```

**New internal helper (not exported):**
```js
// parentId === null  → query top-level me/mailFolders
// parentId === string → query me/mailFolders/{parentId}/childFolders
async function resolveSegmentInParent(accessToken, parentId, segment)
// Returns: Promise<string|null> — folder ID for this segment, or null
```

**Algorithm pseudocode:**
```
getFolderIdByName(accessToken, folderName):
  segments = folderName.split('/').map(s => s.trim()).filter(Boolean)
  if segments.length === 0: return null
  if segments.length === 1: return resolveSegmentInParent(accessToken, null, segments[0])
  currentId = null
  for i, segment in segments:
    currentId = resolveSegmentInParent(accessToken, currentId, segment)
    if currentId === null: return null
  return currentId

resolveSegmentInParent(accessToken, parentId, segment):
  base = parentId ? `me/mailFolders/${parentId}/childFolders` : `me/mailFolders`
  # 1. exact match
  resp = callGraphAPI(accessToken, 'GET', base, null, { $filter: `displayName eq '${segment}'` })
  if resp.value && resp.value.length > 0: return resp.value[0].id
  # 2. case-insensitive fallback
  resp2 = callGraphAPI(accessToken, 'GET', base, null, { $top: 100 })
  if resp2.value:
    match = resp2.value.find(f => f.displayName.toLowerCase() === segment.toLowerCase())
    if match: return match.id
  return null
```

## Graph API Calls

| Level | Endpoint | Query params | When |
|-------|----------|--------------|------|
| Top-level (segment[0]) | `GET me/mailFolders` | `$filter=displayName eq '{seg}'` | exact match, first |
| Top-level fallback | `GET me/mailFolders` | `$top=100` | exact fails |
| Child (segment[n], n>0) | `GET me/mailFolders/{parentId}/childFolders` | `$filter=displayName eq '{seg}'` | exact match, first |
| Child fallback | `GET me/mailFolders/{parentId}/childFolders` | `$top=100` | exact fails |

## Error Handling

| Failure point | Behavior |
|---------------|----------|
| `callGraphAPI` throws (network/401) | caught in `getFolderIdByName` try/catch, returns `null` (existing pattern preserved) |
| Any segment exact match empty + fallback no match | `resolveSegmentInParent` returns `null` → loop returns `null` immediately (short-circuit, no further API calls) |
| Empty path after filtering (`""`, `"/"`, `"//"`) | `segments.length === 0` → return `null` |
| Non-existent top-level | segment[0] returns `null` → loop returns `null`, no child queries made |

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `email/folder-utils.js` | Modify | Add `resolveSegmentInParent()` helper; rewrite `getFolderIdByName()` to split+traverse. `resolveFolderPath`, `getAllFolders`, `WELL_KNOWN_FOLDERS` unchanged. |
| `test/email/folder-utils.test.js` | Modify | Add `describe('getFolderIdByName - path resolution')` block with scenarios below. Existing tests unchanged (backwards compat). |
| `utils/mock-data.js` | Modify | Add child-folder branch in `simulateGraphAPIResponse`: when `path` matches `me/mailFolders/{id}/childFolders`, return mock child folder array keyed by parent id. |

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | `resolveSegmentInParent` exact + fallback (top-level and child) | Mock `callGraphAPI` via `jest.mock`; assert endpoint + queryParams per call |
| Unit | Two-level path `Tramite/REQ-104951` | Chain `mockResolvedValueOnce` per segment; assert leaf id + call count |
| Unit | Three-level path `A/B/C` | 3-segment chain; assert final id |
| Unit | Backwards compat: flat name single API call path | Existing tests must pass unmodified |
| Unit | Case-insensitive segments `tramite/req-104951` | Mock exact-match empty, fallback returns different case; assert id |
| Unit | Non-existent segment returns `null` | Segment[1] exact empty + fallback empty → `null`, short-circuit (no segment[2] calls) |
| Unit | Non-existent top-level returns `null` | Segment[0] empty → `null`, no child queries |
| Unit | Empty segments `Tramite//REQ-104951` | Same mock chain as two-level; assert filtering works |
| Unit | All-empty path `"/"` / `"//"` → `null` | No `callGraphAPI` calls |
| Unit | API error mid-traversal → `null` | Reject on segment[1] call; assert `null` + short-circuit |
| Integration | test-mode server `move-emails` with path | Extend `utils/mock-data.js` child-folder branch; verify via `npm run test-mode` manual smoke |

**Mock data design (`utils/mock-data.js`):** Add a `CHILD_FOLDERS` map keyed by parent id. In the `mailFolders` branch, detect `path.includes('childFolders')` and return the matching child array. Example:
```js
const CHILD_FOLDERS = {
  'tramite-id': [
    { id: 'req-104951-id', displayName: 'REQ-104951' },
    { id: 'req-104952-id', displayName: 'REQ-104952' }
  ]
};
// in mailFolders branch:
if (path.includes('childFolders')) {
  const parentId = path.split('/').find(/* extract id between mailFolders and childFolders */);
  return { value: CHILD_FOLDERS[parentId] || [] };
}
```

## Threat Matrix

N/A — no routing, shell, subprocess, VCS/PR automation, executable-file classification, or process-integration boundary. Change is confined to Graph API folder ID resolution.

## Migration / Rollout

No migration required. Change is backwards compatible (flat names use identical code path). Rollback = `git revert` on `email/folder-utils.js` (single function).

## Open Questions

None — all spec scenarios have a clear implementation path.