# Infrastructure PE Interview Prep Package - MANIFEST

**Generated:** 2026-01-20
**Package Version:** 1.0
**Total Files:** 21 (including this manifest)

---

## Deliverable Index

### Excel Models (5)
| File | Purpose |
|------|---------|
| Model_1_CCGT.xlsx | Combined cycle gas turbine LBO model with 75% cash sweep debt |
| Model_2_Peaker.xlsx | Simple cycle peaker model using dispatch hours (not CF×8760) |
| Model_3_SolarBESS.xlsx | Solar+battery storage model with ITC, degradation, sculpted debt |
| Model_4_Transmission.xlsx | Regulated transmission model with Y0-Y2 construction period |
| Model_5_Midstream.xlsx | Gas gathering model with growth/decline volume profile |

### Drill Versions (5)
| File | Purpose |
|------|---------|
| Drill_1_CCGT.xlsx | CCGT model with Intuition Guide sheet explaining formulas |
| Drill_2_Peaker.xlsx | Peaker model with Intuition Guide sheet explaining formulas |
| Drill_3_SolarBESS.xlsx | Solar+BESS model with Intuition Guide sheet explaining formulas |
| Drill_4_Transmission.xlsx | Transmission model with Intuition Guide sheet explaining formulas |
| Drill_5_Midstream.xlsx | Midstream model with Intuition Guide sheet explaining formulas |

### IC Memos (5)
| File | Purpose |
|------|---------|
| IC_Memo_1_CCGT.docx | Investment Committee memo for CCGT acquisition |
| IC_Memo_2_Peaker.docx | Investment Committee memo for Peaker acquisition |
| IC_Memo_3_SolarBESS.docx | Investment Committee memo for Solar+BESS acquisition |
| IC_Memo_4_Transmission.docx | Investment Committee memo for Transmission acquisition |
| IC_Memo_5_Midstream.docx | Investment Committee memo for Midstream acquisition |

### Reference Materials (4)
| File | Purpose |
|------|---------|
| Master_Template.xlsx | Blank template with structure, formatting, and Instructions sheet |
| Lotus_Interview_CheatSheet_Expanded.md | Quick reference for formulas, conventions, common errors |
| Lotus_Interview_CheatSheet_Expanded.pdf | PDF version of cheat sheet for printing |
| Implementation_Manifesto.md | Detailed explanation of modeling philosophy and decisions |

---

## File Relationships

```
┌─────────────────────────────────────────────────────────────┐
│                    REFERENCE MATERIALS                       │
│  Master_Template.xlsx ──► Starting point for new models     │
│  CheatSheet (md/pdf) ──► Quick formula reference            │
│  Implementation_Manifesto.md ──► Architectural decisions    │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                      MODEL FILES                             │
│  Model_[N]_[Asset].xlsx ──► Complete working LBO models     │
│  - All formulas functional, no Excel errors                 │
│  - Row-index dictionary approach (no hardcoded refs)        │
│  - IB formatting standards (yellow inputs, green outputs)   │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                     DRILL VERSIONS                           │
│  Drill_[N]_[Asset].xlsx ──► Models + Intuition Guide        │
│  - Sheet 1: Same model as Model_[N]                         │
│  - Sheet 2: Intuition Guide explaining each formula         │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                       IC MEMOS                               │
│  IC_Memo_[N]_[Asset].docx ──► Investment recommendation     │
│  - Executive summary, thesis, key metrics                   │
│  - Risk factors, due diligence, IC Q&A                      │
└─────────────────────────────────────────────────────────────┘
```

---

## Technical Standards Applied

All models adhere to conventions documented in `/CLAUDE.md`:
- Row-index dictionary approach (no hardcoded row numbers)
- Freeze panes at E2
- Column layout: A=Labels, B=Units, C=Notes, D=Inputs, E-N=Years 1-10
- Formatting: Yellow inputs (blue bold), green outputs, navy headers
- DSRA targets next year's debt service
- Fuel cost divisor is 1e9 (not 1e6)
- Peakers use dispatch hours (not CF×8760)

---

## Validation Status

All models validated with no Excel errors (#REF, #NAME, #VALUE, #DIV/0).

---

*This manifest is auto-generated. See CLAUDE.md for delivery protocol.*
