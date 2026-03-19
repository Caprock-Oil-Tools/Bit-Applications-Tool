# Plan: Fix Durability Scoring — Radial Primary Cutlet Balance

## Root Cause

The current durability formula uses 8 components that are all **globally aggregated** metrics (overall Gini, overall blade CV, average chamfer fraction, etc.). These partially capture load distribution but miss the **position-specific** signal: **do primary cutlet areas transition smoothly across the radial profile, or are there imbalances at specific radial positions?**

Layout effects (6-3 vulnerability, F-type advantage, redundancy benefits) don't need to be detected by label — they **manifest directly as imbalances in the primary cutlet distribution at specific radial positions**. The SCALE of those imbalances is what determines durability impact.

## What's Missing

At any radial slice, a durable bit has:
- Multiple blades with primary cutlets of **similar area** (balanced load sharing)
- **Smooth transitions** as radius increases (no sudden jumps in cutlet area)

An aggressive/less-durable bit has:
- At certain radii, some blades have **much larger** primary cutlets while others have tiny or no cutlets
- This means fewer blades are doing the real work at those positions = higher force per cutter = less durable
- This is exactly what happens in a 6-3 layout at low IPR: 3 blades have full cutlets, 3 blades are offset and have small/no cutlets

The current `pri_sec_balance` (6% weight) partially captures this as a single global ratio. It needs to be measured **per radial position** and weighted much more heavily.

## New Metrics to Compute (in `compute_cutlet_metrics`)

### 1. `radial_primary_balance` (higher = more durable)
- Bin the radial profile into ~15-20 bins
- For each bin, collect primary (row 1) cutlet areas grouped by blade
- For bins with ≥2 blades represented, compute a balance metric:
  - `1.0 / (1.0 + CV)` where CV = std/mean of per-blade areas in that bin
  - Weight each bin by its share of total primary area (nose/shoulder bins matter more than cone)
- Area-weighted average across all bins = radial_primary_balance
- A 6-3 layout will have low balance at nose/shoulder bins (3 blades large, 3 blades tiny)
- A redundant layout will have high balance (paired blades share similar areas)
- An F-type will have high balance (all blades equally exposed)

### 2. `radial_transition_smoothness` (higher = more durable)
- For each blade, sort its primary cutlets by radial position
- Compute consecutive area ratios: `min(a[i], a[i+1]) / max(a[i], a[i+1])`
- Average ratio across all consecutive pairs across all blades
- Smooth transitions → ratios near 1.0 → high smoothness
- Jumps/discontinuities → ratios near 0 → low smoothness

### 3. `radial_coverage_continuity` (higher = more durable)
- Check what fraction of radial bins have primary cutlets from at least N blades (where N = num_blades/2)
- Gaps in multi-blade coverage = stress concentrations at those positions

## Revised Durability Components & Weights

| Component | Weight | Source | Change |
|---|---|---|---|
| **radial_primary_balance** | 25% | NEW | Key metric: position-specific cross-blade balance |
| **load_balance (Gini)** | 15% | Existing, reduced from 18% | Global load balance still matters |
| **cutter_density** | 15% | Existing, reduced from 22% | Constant volume principle still fundamental |
| **radial_transition_smoothness** | 12% | NEW | Smooth progression = durable |
| **peak_moderation** | 10% | Existing, reduced from 12% | Worst-case cutter overload |
| **chamfer_toughness** | 8% | Existing, reduced from 15% | Chamfer physics still real |
| **blade_balance** | 5% | Existing, reduced from 10% | Partially captured by radial_primary_balance |
| **backrake** | 5% | Existing, reduced from 10% | Conservative angle still matters |
| **backup_engagement** | 5% | Existing, reduced from 7% | Working backups sharing load |

**Why these weights**: The new radial metrics (37% combined) capture the position-specific load distribution that is the PRIMARY driver. The existing global metrics (63%) provide supporting context. `pri_sec_balance` is REMOVED as a separate component because `radial_primary_balance` subsumes it with much finer granularity.

## Implementation Steps

1. **Add new metric computation** in `compute_cutlet_metrics()` (~40 lines):
   - Compute `radial_primary_balance`, `radial_transition_smoothness`, `radial_coverage_continuity`
   - Include them in the returned dict

2. **Update durability scoring** in `main()` (~15 lines):
   - Replace `pri_sec_balance` component with `radial_primary_balance`
   - Add `radial_transition_smoothness` component
   - Rebalance all weights

3. **Update output** to include new metrics in both console and JSON output

4. **Run and verify** results look sensible

## Expected Impact

- Bits with smooth, balanced radial primary distributions → score higher
- Bits with position-specific imbalances (6-3 effect) → score lower, proportional to imbalance severity
- No layout labels used — all signals come from cutlet data
- The "spectrum" nature is preserved: a mild 6-3 offset produces a mild penalty, a severe one produces a severe penalty
