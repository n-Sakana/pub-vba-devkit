# EnvTest Change History (2026-03-29 to 2026-03-30)

## Purpose
This memo records the recent `EnvTest / Survey / Probe` changes, what broke, what was reverted, and what the current trusted state is.

It is intentionally factual. It does not evaluate the design direction. It only records:
- what changed
- which changes damaged reliability
- which commits restored the previous state
- what can be trusted now

## Scope
Relevant files:
- `EnvTest.bat`
- `lib/EnvTest.ps1`
- `lib/internal/Survey.ps1`
- `lib/internal/Probe.ps1`

## Timeline

### 1. Entry-point unification
Commit:
- `7950df4` `Unify environment tests under EnvTest`

Main change:
- unified entry point under `EnvTest`
- moved `Probe.ps1` and `Survey.ps1` under `lib/internal/`
- added `docs/storage-path-strategy.md`

Impact:
- this was the large structural change that made `EnvTest` the main entry
- this commit also added many new investigation items

Notes:
- this commit itself was not the timeout breakage
- later breakage happened on top of this structure

### 2. Prompt order adjustment
Commit:
- `987ae41` `Adjust EnvTest prompt order`

Main change:
- adjusted `EnvTest` prompt order

Impact:
- this became the practical stable baseline before the timeout-related changes

## Reliability breakage

### 3. Probe timeout / execution-path breakage
Commit:
- `3837487` `Tighten EnvTest timeouts and probe verification`

Main change:
- added timeout-related control to `EnvTest`
- changed `Probe` execution path, especially around VBA probe execution

What broke:
- `Probe` no longer preserved the original meaning of VBA-side tests
- `EDR block / reopen failure / run failure / OK` became less trustworthy
- `X (full)` could stop partway through because of outer control behavior

Why this mattered:
- `Probe` is only useful if the executed VBA path is the real one
- changing the wrapper changed the semantics of the result

### 4. Survey timeout refactor
Commit:
- `4323ff8` `Add per-section survey timeouts`

Main change:
- reworked `Survey` collection flow to add timeout behavior

What broke:
- `Survey` changed enough that it could no longer be treated as obviously trustworthy
- timeout behavior was mixed into collection flow before the stable execution path had been preserved

### 5. Partial rollback that was not enough
Commit:
- `7dc9801` `Restore direct VBA probe execution`

Main change:
- removed part of the broken `Probe` wrapper

Problem:
- this was only a partial rollback
- at this point the repo still contained timeout-related churn and trust had not actually been restored

## Recovery

### 6. Restore stable scripts from the pre-timeout baseline
Commit:
- `1a263e7` `Restore stable EnvTest survey and probe scripts`

Main change:
- restored these files to the `987ae41` baseline:
  - `lib/EnvTest.ps1`
  - `lib/internal/Probe.ps1`
  - `lib/internal/Survey.ps1`

Effect:
- removed the broken timeout-driven behavior
- restored the stable execution paths
- re-established a known baseline

This is the key recovery commit.

## Cosmetic-only changes after recovery

### 7. Probe section heading cleanup
Commit:
- `a1145b6` `Reorganize probe sections and report order`

Main change:
- renamed and regrouped `Probe` section headings

Problem introduced:
- report output order was also changed independently of execution order
- that increased the risk of mixing up what ran and what was written

### 8. Restore output order to match execution order
Commit:
- `b181729` `Keep probe report order aligned with execution`

Main change:
- removed the separate result-ordering logic
- restored:
  - `probe.txt` order = execution order

Current meaning:
- `a1145b6` + `b181729` together are cosmetic heading changes only
- there is no separate result reordering anymore

## Current trusted state

Current `main` includes:
- `7950df4`
- `987ae41`
- `1a263e7`
- `a1145b6`
- `b181729`

Current `main` does **not** retain the broken timeout behavior from:
- `3837487`
- `4323ff8`

What can be trusted now:
- `EnvTest` is back on the stable pre-timeout execution path
- `Survey` and `Probe` are not using the broken timeout wrappers
- `probe.txt` output order matches execution order

What is **not** implemented now:
- safe per-item timeout handling
- timeout -> mark only that item -> continue remaining items

That work still needs to be done separately.

## Important lesson from this incident

The failed sequence was:
1. add more investigation items
2. change execution control
3. change timeout behavior
4. change reporting behavior

This mixed too many concerns into the same area.

The required order is:
1. keep stable execution semantics
2. add or reorganize test items
3. only then add timeout handling
4. never separate report order from execution order

## Practical rules for the next change

For `EnvTest / Survey / Probe`, the next implementation should follow these rules:
- result file order must always match execution order
- a timeout must affect only that one item
- timeout must not abort later items
- `OK` must mean the intended action actually completed
- “no exception” alone is not enough for `OK`
- `EDR block`, `run fail`, `timeout`, and `skip` must stay distinct

## Summary

The stable recovery point was restored by:
- `1a263e7`

The current repo keeps only one post-recovery cosmetic change:
- section heading cleanup in `Probe`

The current repo does **not** yet have a correct timeout implementation.
