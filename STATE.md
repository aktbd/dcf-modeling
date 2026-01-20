# Repository State

> **Bridge file for Claude.ai ↔ Claude Code sync.**
> Updated on every significant change.

## Current State

| Field | Value |
|-------|-------|
| Branch | `claude/infrastructure-pe-interview-prep-K9DUF` |
| Latest Tag | `deliverables-20260120-0030` |
| Last Updated | 2026-01-20 02:45 UTC |

## Deliverables

| Asset | Status | Validation |
|-------|--------|------------|
| Model_1_CCGT.xlsx | ✓ Verified | No circular refs, DSRA uses scheduled_ds |
| Model_2_Peaker.xlsx | ✓ Verified | No circular refs, DSRA uses scheduled_ds |
| Model_3_SolarBESS.xlsx | ✓ Verified | No circular refs, sculpted debt (no sweep) |
| Model_4_Transmission.xlsx | ✓ Verified | No circular refs, sculpted debt (no sweep) |
| Model_5_Midstream.xlsx | ✓ Verified | No circular refs, DSRA uses scheduled_ds |
| Drill versions (5) | ✓ Verified | Regenerated with fixed models |
| IC Memos (5) | ✓ Valid | DOCX format |
| Reference Materials | ✓ Valid | Template, CheatSheet, Manifesto |

**Download:** [deliverables-20260120-0030.zip](https://github.com/aktbd/dcf-modeling/raw/claude/infrastructure-pe-interview-prep-K9DUF/deliverables/deliverables-20260120-0030.zip)

## Formula Verification (2026-01-20 02:15)

**Fix Applied:** DSRA Target circular reference resolved in cash sweep models (CCGT, Peaker, Midstream)

- **Root cause:** DSRA Target Y(n) referenced debt_service_pre Y(n+1), which depended on ending balance, which depended on cash sweep, which depended on DSRA funding
- **Fix:** Added `scheduled_ds` row using straight-line amortization only (independent of sweep)
- **Sculpted models** (Solar+BESS, Transmission): No fix needed - debt service determined by CFADS/target_DSCR, independent of DSRA

**Verification Method:** `formulas` Python library for formula evaluation

- All models load and calculate without circular reference errors
- No #REF, #NAME, #VALUE, #DIV/0 errors detected
- DSCR values in expected ranges (1.25x-1.40x)
- Fuel cost uses 1e9 divisor (correct)

## Recent Activity

- 2026-01-20 02:45: Formula cleanup - replaced /1000000000 with /1E9 for readability
- 2026-01-20 02:35: Confirmed GitHub URLs as primary delivery (sandbox network blocked)
- 2026-01-20 02:30: Added retrospective, Drive status, skills audit to CLAUDE.md v1.5
- 2026-01-20 02:15: Fixed circular reference in DSRA, regenerated all models and drills
- 2026-01-20 02:00: Initial formula verification identified circular reference issue
- 2026-01-20 01:45: Cleaned up delivery protocol
- 2026-01-20 00:30: Initial deliverables published

## Known Issues

None.

## Quick Links for Claude.ai

| Resource | Raw URL |
|----------|---------|
| This file | `https://raw.githubusercontent.com/aktbd/dcf-modeling/claude/infrastructure-pe-interview-prep-K9DUF/STATE.md` |
| Index | `https://raw.githubusercontent.com/aktbd/dcf-modeling/claude/infrastructure-pe-interview-prep-K9DUF/OUTPUTS_INDEX.md` |
| Standards | `https://raw.githubusercontent.com/aktbd/dcf-modeling/claude/infrastructure-pe-interview-prep-K9DUF/CLAUDE.md` |
