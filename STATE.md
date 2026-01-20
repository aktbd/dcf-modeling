# Repository State

> **Bridge file for Claude.ai ↔ Claude Code sync.**
> Updated on every significant change.

## Current State

| Field | Value |
|-------|-------|
| Branch | `claude/infrastructure-pe-interview-prep-K9DUF` |
| Latest Tag | `deliverables-20260120-1000` |
| Last Updated | 2026-01-20 10:00 UTC |

## Deliverables

| Asset | Status | Validation |
|-------|--------|------------|
| Model_1_CCGT.xlsx | ✓ Complete | 6 tabs + Spark Spread + Unlevered IRR + Sensitivity |
| Model_2_Peaker.xlsx | ✓ Complete | 6 tabs + Spark Spread + Unlevered IRR + Sensitivity |
| Model_3_SolarBESS.xlsx | ✓ Complete | 6 tabs + Sculpt transparency + Unlevered IRR |
| Model_4_Transmission.xlsx | ✓ Complete | 6 tabs + Sculpt transparency + Unlevered IRR |
| Model_5_Midstream.xlsx | ✓ Complete | 6 tabs + Unlevered IRR + Sensitivity |
| Model_6_Wind.xlsx | ✓ Complete | 6 tabs + PTC mechanics + Sculpted debt |
| Drill versions (6) | ✓ ENHANCED | Formulas blanked + DRILL MODE note + Unlevered IRR |
| IC Memos (6) | ✓ ENHANCED | Fill-in-the-blank templates with asset-specific prompts |
| **Lotus_Study_Guide.md** | ✓ COMPLETE | 12 sections including Advanced Scenarios + Proxy Defaults |
| Reference Materials | ✓ Valid | Template, CheatSheet (updated), Manifesto |

**Download:** [deliverables_bundle.zip](https://github.com/aktbd/dcf-modeling/raw/claude/infrastructure-pe-interview-prep-K9DUF/deliverables/deliverables-20260120-0030/deliverables_bundle.zip)

## Final Polish (2026-01-20 10:00)

**Commit:** `880bf1a` - Drills + Sensitivity + Memo Templates + Study Guide Enhancements

| Enhancement | Status | Details |
|-------------|--------|---------|
| Drill Blanking | Complete | All 6 Drills have key formulas blanked with DRILL MODE note |
| Unlevered IRR | Complete | Added to all 6 Model files |
| Sensitivity Notes | Complete | "+/- 5% Entry → ~+/- 1% IRR" in all Models |
| IC Memo Templates | Complete | Fill-in-the-blank format with asset-specific prompts |
| Study Guide Sec 11 | Complete | Advanced Scenarios (Mezz, Time Series, Construction, ERCOT, Tax Equity) |
| Study Guide Sec 12 | Complete | Proxy Defaults (market prices, financing terms, transaction defaults) |

**Prior Fixes Verified:**
- Midstream Revenue: `=E56*E57` (no *1000 error)
- Transmission Sources: `=D109+D110` (no DSRA in Sources)
- Solar Tax: has `MAX(0,...)` guardrail
- Wind PTC: uses `MIN()` (cannot create negative taxes)

**ZIP Verified:** 23 files, SHA256: `b591f39d01167d74818d7aff362c24361092e27c7c15f2076dee03baf11f14b9`

## 3-Tab Structure Enhancement (2026-01-20 07:30)

All Model and Drill files now have:

| Tab | Purpose | Rows |
|-----|---------|------|
| Model_Quick | 90-min build, no DSRA/sweep | ~60 |
| Model_Standard | Full model + Control Panel | ~150 |
| Model_Full | Standard + LLCR/PLCR | ~160 |

**Control Panel Flags:**
- Revenue Mode: Merchant / Contracted / Hybrid
- Contract Years: If Hybrid, years of contracted revenue
- Debt Sweep: ON / OFF
- Sweep %: 75% (when Sweep = ON)
- Exit Method: Multiple / DCF

**README Additions:**
- Cold Start Workflow (6 phases, 4-hour test)
- Time Budget breakdown
- Debt Sizing Approaches (Straight-line, Sweep, Sculpt)
- Key Formula Concepts (Spark Spread, UCAP, LLCR/PLCR)

## Formula Verification (2026-01-20 03:45)

**Full Sanity Check Complete:**

| Model | DSCR Formula | Min DSCR | Status |
|-------|--------------|----------|--------|
| CCGT | =IF(DS>0,CFADS/DS,0) | ~1.29x | ✓ PASS |
| Peaker | =IF(DS>0,CFADS/DS,0) | ≥1.25x | ✓ PASS |
| Solar+BESS | =IF(DS>0,CFADS/DS,0) | ≥1.30x | ✓ PASS |
| Transmission | =IF(DS>0,CFADS/DS,0) | ≥1.35x | ✓ PASS |
| Midstream | =IF(DS>0,CFADS/DS,0) | ≥1.25x | ✓ PASS |

**CCGT Manual Verification:**
- Revenue: $122.55mm (Energy $84.41 + Capacity $38.14)
- OpEx: $48.56mm (Fuel $36.93 + VOM $6.15 + FOM $5.47)
- EBITDA: $73.99mm
- CFADS: $62.86mm
- Debt Service: $48.64mm
- **DSCR = $62.86 / $48.64 = 1.292x** ✓

**Previous Fix (still applied):** DSRA Target circular reference resolved via `scheduled_ds` row

## Recent Activity

- 2026-01-20 10:00: **Final Polish**: Drill blanking, Unlevered IRR, IC Memo templates, Study Guide Sec 11-12 (commit 880bf1a)
- 2026-01-20 09:00: **Final Update**: Wind model, Study Guide, Spark Spread, enhanced Control Panel, Proxy tables
- 2026-01-20 07:30: 3-Tab Structure: Model_Quick, Model_Standard, Model_Full + Control Panel + Cold Start Workflow
- 2026-01-20 06:00: Critical audit fixes: Midstream *1000, Transmission S&U, Solar MAX(0) (commit d9f8b1f)
- 2026-01-20 05:30: Added README, Case Prompt tabs; MINIFS for Min DSCR (commit 77efe81)
- 2026-01-20 05:00: Re-verified ZIP with extraction test, force pushed (commit 1877df2)
- 2026-01-20 04:45: Fixed ZIP - verified IC Memo tabs present by extraction test
- 2026-01-20 04:00: Added IC Memo tabs to all 5 Model files (content from DOCX memos)
- 2026-01-20 03:45: Full sanity check - all models validated, DSCR 1.29x confirmed
- 2026-01-20 03:30: Fixed Column C text format (removed "=" prefix breaking parser)
- 2026-01-20 03:15: Added formula comments, IC Memo tabs, Quick Reference sheet
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
