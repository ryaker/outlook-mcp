# Tasks: Fix `move-emails` to support subfolder targets

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~270 |
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
| 1 | Full change (RED→GREEN→REFACTOR) | PR 1 | `npm test -- test/email/folder-utils.test.js` | `npm run test-mode` + manual `move-emails` with path | Revert `email/folder-utils.js` only |

## Phase 1: RED — Write Failing Tests

- [x] 1.1 Add `CHILD_FOLDERS` mock map and child-folder branch in `utils/mock-data.js` `simulateGraphAPIResponse()`
- [x] 1.2 Add `describe('getFolderIdByName - path resolution')` block in `test/email/folder-utils.test.js` with 11 test cases covering: two-level path, three-level path, case-insensitive segments, non-existent segment, non-existent top-level, empty segments `Tramite//REQ-104951`, all-empty path `"/"`/`"//"`, API error mid-traversal, and `resolveSegmentInParent` exact+fallback (top-level and child)
- [x] 1.3 Run `npm test` — confirm new tests FAIL (function not yet path-aware)

## Phase 2: GREEN — Implement Path Resolution

- [x] 2.1 Add `resolveSegmentInParent(accessToken, parentId, segment)` helper in `email/folder-utils.js` with exact-match + case-insensitive fallback, querying `me/mailFolders` (parentId=null) or `me/mailFolders/{parentId}/childFolders` (parentId set)
- [x] 2.2 Rewrite `getFolderIdByName()` to split `folderName` on `/`, filter empty segments, traverse via `resolveSegmentInParent` loop, short-circuit on `null`
- [x] 2.3 Run `npm test` — confirm all 11 new tests PASS and all existing tests remain green

## Phase 3: REFACTOR & Verify

- [x] 3.1 Run full `npm test` suite — verify zero regressions across all modules
- [x] 3.2 Verify backwards compatibility: existing `resolveFolderPath` and `getFolderIdByName` tests unchanged, flat names use identical code path
