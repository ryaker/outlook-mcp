## Exploration: Fix `move-emails` to support subfolder targets

### Current State

`move-emails` accepts a `targetFolder` parameter as a flat folder name string (e.g. `"Tramite"`). It calls `getFolderIdByName()` which queries `me/mailFolders?$filter=displayName eq '{name}'` — this only searches **top-level** folders. There is no path resolution, no hierarchy traversal, and no `/` separator logic.

The same `getFolderIdByName()` is used by **four call sites**: `folder/move.js`, `folder/create.js` (twice, for existence check and parent lookup), and `rules/create.js`. None of them currently support subfolder paths.

Helper functions already exist for hierarchy traversal:
- `getAllFolders()` in `email/folder-utils.js` (lines 120-168) fetches top-level folders + one level of `childFolders`, but it returns a **flat list** with no path structure, and only goes **one level deep**.
- `getAllFoldersHierarchy()` in `folder/list.js` (lines 65-129) does the same one-level fetch with additional formatting.
- `create-folder` already has a `parentFolder` parameter and resolves it via `getFolderIdByName(accessToken, parentFolderName)` at `folder/create.js:79` — but **same limitation**: parentFolder must be top-level.

The `find-folder-ids.js` utility script (lines 29-82) demonstrates the correct Graph API pattern: fetch top-level folders, find the parent by name, then query `me/mailFolders/{parentId}/childFolders` to get children. This is the pattern to follow.

### Affected Areas

- `email/folder-utils.js` — Add `resolveFolderPathById()` and/or enhance `getFolderIdByName()` with path parsing logic
- `folder/move.js` — Update `moveEmailsToFolder()` to handle path-style `targetFolder` (e.g. `"Tramite/REQ-104951"`)
- `folder/index.js` — No schema changes needed unless adding a `parentFolder` param (not recommended)
- `folder/create.js` — Already has `parentFolder`; could benefit from path resolution but **out of scope** for this change
- `rules/create.js` — Caller of `getFolderIdByName()`; **out of scope** but should be noted as future beneficiary
- `email/search.js`, `email/list.js` — Callers of `resolveFolderPath()`; could benefit from subfolder support but **out of scope**
- `utils/mock-data.js` — May need mock data for path-based resolution if tests require it
- `test/email/folder-utils.test.js` — Tests for `getFolderIdByName` and `resolveFolderPath` will need new test cases

### Approaches

1. **Enhance `getFolderIdByName()` with path splitting** — Add path-aware resolution directly in the core lookup function
   - Pros: Single change point fixes ALL callers automatically; trivial to follow the `find-folder-ids.js` pattern
   - Cons: Changes behavior of a shared utility; all callers get subfolder support even if they don't need it (acceptable)
   - Effort: **Low**

2. **Create a new `resolveFolderPathById()` utility** — Keep `getFolderIdByName()` unchanged; add a new function specifically for path resolution
   - Pros: Clean separation of concerns; no blast radius to other callers
   - Cons: `folder/move.js` must call new function; `folder/create.js` would need separate update for subfolder parent support
   - Effort: **Low**

3. **Replace `getFolderIdByName()` entirely** — Rewrite to always traverse the full hierarchy
   - Pros: Most thorough; handles nested paths of arbitrary depth
   - Cons: Performance cost from multiple API calls for flat lookups; existing callers that only need top-level search pay unnecessary cost
   - Effort: **Medium**

4. **Add `parentFolder` parameter to `move-emails`** (like `create-folder` has) — Keep flat `targetFolder` but add an optional `parentFolder` field
   - Pros: Parallels the existing `create-folder` pattern; most explicit for users
   - Cons: Doesn't solve the original problem (user asks for `"Tramite/REQ-104951"`); requires documentation changes; two params where one path string would suffice
   - Effort: **Low**

### Recommendation

**Approach 1: Enhance `getFolderIdByName()` with path splitting.**

Rationale:
- Straightforward implementation: split input on `/`, trim segments, traverse `childFolders` level by level using the Graph API `me/mailFolders/{parentId}/childFolders` endpoint
- Backwards compatible: if input has no `/`, behavior is identical to today (exact match via `$filter`, fallback to case-insensitive)
- All four call sites benefit automatically
- Follows the proven pattern from `find-folder-ids.js`
- The `getAllFolders()` and `getAllFoldersHierarchy()` functions already prove the `childFolders` endpoint works

The implementation would:
1. Check if `folderName` contains `/`; if not, use existing logic (zero blast radius)
2. Split on `/`, trim each segment
3. Resolve the first segment via existing top-level lookup (`$filter` on `displayName`)
4. For each subsequent segment, query `me/mailFolders/{parentId}/childFolders?$filter=displayName eq '{segment}'`
5. Return the final folder ID or `null` if any segment fails to resolve

### Risks

- **Performance**: Deeply nested paths (e.g. 5+ levels) could add latency from sequential API calls. Mitigation: this is a per-request operation and typical mailbox folder depth is 2-3 levels.
- **Folder names with `/`**: If a folder literally contains `/` in its display name, path splitting will misparse. Mitigation: extremely rare in practice; can be documented as a known limitation.
- **Case sensitivity**: The Graph API `$filter` is case-sensitive by default. The existing code already handles this (exact match then case-insensitive fallback) but the childFolder endpoint may behave differently. The fallback should fetch child folders and match case-insensitively.
- **Cache semantics**: `folderCache` exists at `email/folder-utils.js:10` but is never populated or used — dead code. Don't reintroduce caching without a clear invalidation strategy.
- **Pagination**: A parent could have >100 child folders (unlikely for mail folders, but possible). The child folder query should use `$top=100` and potentially paginate.

### Ready for Proposal

**Yes.** The problem is well-understood, the fix is scoped, and Approach 1 is the clearest path forward. The orchestrator should proceed with `sdd-propose` for the `fix-subfolder-move-emails` change.

Key points to tell the user:
- Only `email/folder-utils.js` needs substantive changes
- No schema changes to `move-emails` tool definition
- All existing callers work unchanged (backwards compatible)
- Blast radius is contained — only `getFolderIdByName()` is modified
- Edge case of literal `/` in folder names is a known limitation
