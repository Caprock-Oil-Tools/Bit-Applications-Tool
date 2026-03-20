# AI Learned Notes — Layout Durability Insights

## Source: Comparative Cutlet Analysis (March 2026)

Studied three bit designs at three IPR values each (0.100, 0.250, 0.500 in/rev):
- **8896r00** — F-type, 0.100" 6-3 shift (secondary blades under-exposed by 0.100")
- **2746r00** — F-type, 0.035" 6-3 shift (secondary blades under-exposed by 0.035")
- **2226r1** — True 6-3 layout (secondary blades only in shoulder/gauge zone)

---

## Key Finding 1: Imbalance Is a Continuous Spectrum

The primary-to-secondary blade cutlet area ratio varies continuously:

| IPR | 2746 (0.035" shift) | 8896 (0.100" shift) | 2226 (6-3) |
|---|---|---|---|
| 0.100 | 8.2:1 | 3.0:1 | 429:1 |
| 0.250 | 3.9:1 | 3.0:1 | 16.5:1 |
| 0.500 | 3.3:1 | 2.5:1 | 8.3:1 |

There is no clean binary between "F-type" and "6-3." They sit on a curve.
The durability score must measure WHERE on that curve a design sits, not
just classify it into a bucket.

---

## Key Finding 2: Smaller Shift Does NOT Mean Less Imbalance

At IPR=0.100:
- 2746 (0.035" shift) = 8.2:1 ratio — WORSE than 8896
- 8896 (0.100" shift) = 3.0:1 ratio — BETTER balance

Why: A very small shift puts secondary cutters almost directly behind
primaries. At thin cut depths (low IPR), they barely peek out. A larger
shift gives secondaries more independent radial exposure.

The relationship is non-linear: it's the shift-to-IPR ratio that determines
imbalance, not the shift alone.

---

## Key Finding 3: All Designs Converge at High IPR

At IPR=0.500, even the 6-3 drops to 8.3:1. Deep cuts overwhelm the
shift geometry. This means:
- Durability penalties should be most sensitive at LOW IPR
- The low-IPR ratio is the "worst case" that limits bit life
- A bit that's only balanced at high IPR is vulnerable when the driller
  reduces WOB (which happens often — sliding, connection breaks, etc.)

---

## Key Finding 4: The Low-IPR Ratio Is the Right Durability Metric

Computing cutlets at IPR=0.100 and measuring the primary/secondary ratio
naturally captures:
- The shift geometry (without needing to know the shift amount)
- Whether secondary blades truly engage or just nominally exist
- The "worst case" imbalance that limits bit life

At IPR=0.100:
- Redundant/balanced design: ratio near 1:1
- F-type with good shift: ratio 2-4:1 (proportional, all blades working)
- F-type with tiny shift: ratio 5-10:1 (secondaries starved)
- True 6-3: ratio 100+:1 (secondaries essentially absent)

---

## Key Finding 5: What "Proportional Imbalance" Means (User Input)

The user described the F-type as a "proportional imbalance":
- Primary blades consistently do ~2-3x the work of secondary blades
- This ratio is roughly CONSTANT across the profile (not random)
- All blades ARE doing work at every IPR — no zero-engagement blades
- The imbalance is structured and predictable

This is fundamentally different from a 6-3 layout where:
- Secondary blades may have ZERO engagement at low IPR
- The ratio changes dramatically with IPR (429:1 -> 8.3:1)
- The imbalance is unstable and IPR-dependent

---

## Key Finding 6: The Previous Scoring Was Under-Weighting Layout Balance

The prior `pri_sec_balance` component had only 6% weight in durability.
This was insufficient because:
1. A 6-3 at low IPR has 3 blades doing zero work — a 50% reduction in
   effective blade count, catastrophic for durability
2. An F-type with bad shift can be nearly as bad (8.2:1 at low IPR)
3. The pri_sec_ratio was only computed at the file's default IPR (0.250),
   missing the worst-case low-IPR behavior entirely

The fix: compute cutlets at IPR=0.100 in addition to file IPR, and weight
the low-IPR balance heavily (20%). This captures the F-type vs 6-3
distinction, the shift magnitude effect, and IPR sensitivity — all from
the physics.

---

## Improved Durability Scoring (v2)

### New components and weights:
| Component | Weight | Description |
|---|---|---|
| low_ipr_layout_balance | 20% | 1/log2(ratio+1) at IPR=0.100 |
| cutter_density | 15% | Cutlets per sq inch of bit face |
| load_balance | 12% | 1 - Gini coefficient |
| chamfer_toughness | 12% | Average chamfer fraction |
| backrake | 10% | avg_backrake / 30 |
| peak_moderation | 8% | 1 / max_mean_ratio |
| blade_balance | 8% | 1 / (1 + blade_cv) |
| backup_engagement | 7% | Backup cutlet fraction |
| layout_stability | 8% | 1 / (low_ratio / file_ratio) |

### Result on study bits:
| Bit | Old Dur | New Dur | PS@Low IPR | Description |
|---|---|---|---|---|
| 8896 | 1.5 | 4.5 | 0.7:1 | F-type, good shift — best balance |
| 2746 | 3.5 | 2.5 | 2.2:1 | F-type, small shift — mediocre |
| 2226 | 0.4 | 0.7 | 3.8:1 | True 6-3 — worst layout balance |

---

## User-Provided Design Context

From user inputs during this study:
- "Think of the F-type as a proportional imbalance"
- "8896 has a 0.100" 6-3 shift under-exposed on the secondary blades"
- "2746 has a 0.035" 6-3 shift under-exposed on the secondary blades"
- "Even though these are different designs you can see that the amount of shift matters"
- The shift amount is a design parameter that varies continuously
- Both F-type and 6-3 are on the same spectrum, just at different points
