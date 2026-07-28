# Proposal: Fix `move-emails` to support subfolder targets

## Intent

`move-emails` fails when the user specifies a subfolder path like `"Tramite/REQ-104951"` because `getFolderIdByName()` only searches top-level folders via `me/mailFolders?$filter=displayName eq '{name}'`. Users managing deeply nested folder hierarchies (e.g., Tramite → REQ-104951) cannot move emails to subfolders without manually looking up folder IDs.

## Scope

### In Scope
- Enhance `getFolderIdByName()` in `email/folder-utils.js` to resolve path-style folder names (e.g., `"Tramite/REQ-104951"`)
- All 4 call sites benefit automatically: `folder/move.js`, `folder/create.js` (x2), `rules/create.js`
- New test cases in `test/email/folder-utils.test.js` for path resolution
- Backwards compatible: flat folder names (no `/`) behave identically

### Out of Scope
- `create-folder` subfolder parent support via path (deferred)
- `rules/create.js` subfolder support for `moveToFolder` (deferred)
- `email/search.js` / `email/list.js` subfolder support (deferred)
- Folder names with literal `/` in display name (known limitation)
- Caching layer (`folderCache` is dead code — not reintroduced)

## Capabilities

### New Capabilities
None — this is a behavioral enhancement to an existing utility function.

### Modified Capabilities
None — no spec-level behavior changes. The `move-emails` tool contract (accepts `targetFolder` string) is unchanged. Only internal resolution logic improves.

## Approach

**Enhance `getFolderIdByName()` with path splitting** (Approach 1 from exploration):

1. Check if `folderName` contains `/`; if not, use existing logic (zero blast radius)
2. Split on `/`, trim each segment
3. Resolve first segment via existing top-level `$filter` lookup
4. For each subsequent segment, query `me/mailFolders/{parentId}/childFolders?$filter=displayName eq '{segment}'`
5. Apply case-insensitive fallback at each level (same pattern as current code)
6. Return final folder ID or `null` if any segment fails

Follows the proven traversal pattern from `find-folder-ids.js`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `email/folder-utils.js` | Modified | `getFolderIdByName()` gains path-splitting logic |
| `test/email/folder-utils.test.js` | Modified | New test cases for path resolution |
| `utils/mock-data.js` | Modified | Mock data for child folder API responses |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Deep nesting adds latency from sequential API calls | Low | Typical mailbox depth is 2-3 levels; per-request operation |
| Folder names with literal `/` misparsed | Low | Document as known limitation; extremely rare in practice |
| Child folder pagination (>100 children) | Low | Use `$top=100`; pagination unlikely for mail folders |

## Rollback Plan

Revert the single function modification in `email/folder-utils.js`. The change is contained to one function — no schema changes, no tool definition changes. Simple `git revert` on the affected file.

## Dependencies

- None (pure enhancement to existing utility)

## Success Criteria

- [ ] `move-emails` with `targetFolder: "Tramite/REQ-104951"` resolves the subfolder and moves emails
- [ ] `move-emails` with `targetFolder: "Tramite"` (flat) continues to work identically
- [ ] All existing tests pass without modification
- [ ] New tests cover: single-segment (backwards compat), two-segment path, three-segment path, non-existent segment, case-insensitive segments, empty path
