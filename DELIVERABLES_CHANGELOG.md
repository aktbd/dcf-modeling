# Deliverables Changelog

Append-only history of deliverable drops.

---

## deliverables-20260120-0030

**Date:** 2026-01-20 00:30 UTC

**What:** Initial drop - Infrastructure PE Interview Prep Package
- 5 Excel LBO models (CCGT, Peaker, Solar+BESS, Transmission, Midstream)
- 5 Drill versions with Intuition Guide sheets
- 5 IC Memos (Word documents)
- Reference materials (Template, CheatSheet, Manifesto)

**Location:** [`/deliverables/deliverables-20260120-0030/`](./deliverables/deliverables-20260120-0030/)

**ZIP sha256:** `0de34c32c1652ac71b15cc35231b640002d07634fccda0aafa2dfbd67d11cee4`

### Update 2026-01-20 02:15 UTC (in-place fix, same tag)

**Bug Fix:** Circular reference in DSRA Target formula (cash sweep models)

- **Affected models:** CCGT, Peaker, Midstream (cash sweep structure)
- **Root cause:** DSRA Target referenced debt_service_pre which depended on cash sweep, creating circular dependency
- **Fix:** Added `scheduled_ds` row using straight-line amortization to break circularity
- **Verification:** All models pass formula evaluation using `formulas` Python library
- **Not affected:** Solar+BESS, Transmission (sculpted debt, no circularity)

---
