# Archive Report: fix-subfolder-move-emails

- **Archived at**: 2026-07-28
- **Artifact store**: hybrid (openspec + engram)
- **Evidence revision**: `sha256:f5f59f09a14b8ad9c6ae8c5549ddfff6d5f00b71ad9cd7bb5a0aa0a5eabdf0ce`

## Verdict

| Metric | Value |
|--------|-------|
| Verdict | **PASS** |
| Blockers | 0 |
| Critical findings | 0 |
| Requirements compliant | 6/6 |
| Scenarios compliant | 8/8 |
| Tests passing | 142/142 |
| Tasks complete | 8/8 |

## Artifact Lineage (Engram Observation IDs)

| Artifact | Engram ID | Sync ID |
|----------|-----------|---------|
| Exploration | #1866 | `obs-8564a2e4f3c46b78` |
| Proposal | #1868 | `obs-f5c33f9ac81cc17c` |
| Spec (delta) | #1869 | `obs-0e54b0741a45483f` |
| Design | #1870 | `obs-29b8b4c300451763` |
| Tasks | #1871 | `obs-94dec7d0626df1c1` |
| Apply progress | #1872 | `obs-9cb00a6a6ebb2b91` |
| Verify report | #1874 | `obs-edb037256de98ec5` |
| Archive report | — | — |

## Specs Synced

| Domain | Action | Details |
|--------|--------|---------|
| email | Created (new main spec) | 7 requirements, 8 scenarios — delta spec copied as full main spec since `openspec/specs/email/spec.md` did not exist |

## Summary

The change enhanced `getFolderIdByName()` in `email/folder-utils.js` to resolve path-style folder names (e.g., `"Tramite/REQ-104951"`) by splitting on `/` and traversing the folder hierarchy level-by-level via `resolveSegmentInParent()` helper. All 4 call sites (`folder/move.js`, `folder/create.js` x2, `rules/create.js`) benefit transparently — no schema or tool definition changes.

### Design Deviations

- `resolveSegmentInParent` was **exported** from `email/folder-utils.js` (design said "not exported") for direct testability per tasks requirement. Documented and accepted during apply phase.
- Two mock call-count assertions (`toHaveBeenCalledTimes`) used to verify short-circuit behavior — acceptable for the purpose.

## Final State

- **142 tests passing** across 7 test suites (verified at archive time — zero regressions)
- **0 warnings** carried forward
- **0 open issues**
- `openspec/specs/email/spec.md` now reflects the new behavior as the source of truth
- Archive located at `openspec/changes/archive/2026-07-28-fix-subfolder-move-emails/`
