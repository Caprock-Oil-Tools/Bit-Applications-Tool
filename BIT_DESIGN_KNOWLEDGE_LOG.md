# PDC Bit Design Knowledge Log

## Reference for all design parameters, scoring logic, and domain insights

> **Purpose**: This file captures every bit design insight, rule, and relationship
> provided by the user and extracted from project documents. Consult this file
> whenever evaluating aggressiveness, durability, steerability, or any other
> design characteristic of a PDC drill bit.

---

## TABLE OF CONTENTS

1. [Glossary & Abbreviations](#1-glossary--abbreviations)
2. [Bit Profile Zones](#2-bit-profile-zones)
3. [Layout Types & Force Characteristics](#3-layout-types--force-characteristics)
4. [Ripper / Cutter Configurations](#4-ripper--cutter-configurations)
5. [Backup Element Placement & Engagement](#5-backup-element-placement--engagement)
6. [PDC-Rock Interaction Fundamentals](#6-pdc-rock-interaction-fundamentals)
7. [Helical Path of Motion (POM) & IPR](#7-helical-path-of-motion-pom--ipr)
8. [Cutlet Plots & Force Calculations](#8-cutlet-plots--force-calculations)
9. [Durability Scoring Model](#9-durability-scoring-model)
10. [Steerability Scoring Model](#10-steerability-scoring-model)
11. [Key Metrics Definitions](#11-key-metrics-definitions)
12. [Design Trade-offs & Rules of Thumb](#12-design-trade-offs--rules-of-thumb)
13. [Min Engagement File Structure](#13-min-engagement-file-structure)
14. [Cutter Naming Convention](#14-cutter-naming-convention)
15. [Data Sources & File Inventory](#15-data-sources--file-inventory)

---

## 1. Glossary & Abbreviations

| Term | Definition |
|------|-----------|
| **BHA** | Bottom Hole Assembly — the assembly of tools at the bottom of the drillstring |
| **Backrake** | Angle of the cutter face relative to the cutting direction; higher = more conservative/durable, lower = more aggressive |
| **Chamfer** | Beveled edge of a PDC that engages rock first at lower DOC; increases compression on the diamond |
| **Cutlet** | The portion of a PDC or knuckle element that does work at a given IPR; its area and centroid determine the volume of rock each element interacts with |
| **DBR** | Differential Bit Rotation — catastrophic failure mode where part of the bit rotates at a different speed; prevented by properly engaged backup cutters that signal wear |
| **DOC** | Depth of Cut — how deep each cutter penetrates per revolution |
| **Dragon Layout** | Layout where trailing backup cutters have their own blades, preventing pack-off and maintaining ROP |
| **Element Vector** | Spatial orientation of a cutter combining tilt and backrake; determines the direction forces are applied to each cutlet |
| **Engagement Magnitude** | Volume and force characteristics of how much a cutter interacts with rock |
| **Engagement Rate** | Speed at which engagement intensity increases with changes in drilling parameters |
| **Engagement Speed (Min)** | The minimum IPR (or ft/hr at a given RPM) at which a secondary element begins doing work |
| **F-Type** | Layout where secondary blades have radially shifted cutters; similar aggressiveness to 6-3 but all blades equally exposed = more durable at low IPR |
| **Forward Spiral** | Layout (blade order 1-2-3-4-5-6) where cutlet forces are predominantly radial, compressing the bit body; requires more torque |
| **IPR** | Inches Per Revolution = ft/hr x 12 / 60 / RPM. Alternative: ft/hr x rpm x 0.2 |
| **JSA** | Junk Slot Area — space between blades for cuttings evacuation |
| **Knuckle** | Non-PDC backup element (often CPS type) used to limit ROP and improve steerability; engages differently than PDC backups |
| **Mixed Cutlet Type** | Layout employing all three cutlet shapes (Forward, Reverse, Redundant) in one design; multi-directional rock attack |
| **Pack-off** | Accumulation of cuttings around underexposed trailing cutters that lack space to evacuate; slows ROP |
| **PDC** | Polycrystalline Diamond Compact — the primary cutting element |
| **POM** | Path of Motion — the helical path a cutter traces during drilling |
| **Redundant Layout** | Layout (blade pairs 14-25-36 or 14-36-25) where cutlet forces push perpendicular to the bit profile; general purpose, durable |
| **Reverse Spiral** | Layout (blade order 1-6-5-4-3-2) where cutlet forces are predominantly axial; requires more WOB; used for curves and laterals |
| **Ripper** | Specialty cutter shape (not full-round) that is more aggressive than standard PDC |
| **ROP** | Rate of Penetration — drilling speed in ft/hr |
| **Siderake** | Lateral angle of the cutter; creates a directional force component |
| **Tilt** | Angular component of cutter orientation combined with backrake to form the element vector |
| **WOB** | Weight on Bit — downward force applied to the bit |

---

## 2. Bit Profile Zones

The bit face is divided into radial zones measured as fractions of gauge radius:

| Zone | Radial Range | Characteristics |
|------|-------------|-----------------|
| **Cone** | 0% – 30% of gauge radius | Center of bit; low linear velocity; typically lower backrake |
| **Nose** | 30% – 60% of gauge radius | Transition zone; highest cutting volume; critical for ROP |
| **Taper** | 60% – 85% of gauge radius | Transition to gauge; important for directional control |
| **Gauge** | 85% – 100% of gauge radius | Outer edge; determines hole diameter; critical for steerability |

**Zone engagement columns in BAT workbook (BA-BE):**
- BA: 6-3 engage — min ft/hr@100rpm for primary PDCs on secondary blades to engage
- BB: Nose PDC engage — min ft/hr@100rpm for secondary PDC backups at nose
- BC: Taper PDC engage — min ft/hr@100rpm for secondary PDC backups at taper
- BD: Nose knuckle engage — min ft/hr@100rpm for knuckle backups at nose
- BE: Taper knuckle engage — min ft/hr@100rpm for knuckle backups at taper

---

## 3. Layout Types & Force Characteristics

### Redundant (14-25-36 or 14-36-25)
- **Forces**: Perpendicular to bit profile
- **Application**: General purpose, durable
- **Detected by**: High blade-pair radial overlap (redundancy_score)
- **Blade exposure**: Secondary blades intentionally cover less of the profile — this is a feature for backup, not a flaw

### Forward Spiral (1-2-3-4-5-6)
- **Forces**: Predominantly radial; compress the bit body
- **Requires**: More torque
- **Application**: Surface bits, smaller rigs with limited WOB wanting higher ROP
- **Trade-off**: Higher ROP but increased torque demand

### Reverse Spiral (1-6-5-4-3-2)
- **Forces**: Predominantly axial
- **Requires**: More WOB
- **Application**: Curve and lateral sections for higher overall ROP
- **Advantage**: Increased tool face control = better steerability

### 6-3 Layout
- **Can be applied to any bit design**
- **Effect**: Makes any design more aggressive
- **Vulnerability**: At low IPR, effectively functions as a three-bladed bit (only 3 of 6 blades doing work)
- **Detected by**: Z-position differences between blade groups at the nose (six_three_offset metric)

### F-Type (4+ variants)
- **Similar aggressiveness to 6-3** but all blades equally exposed
- **More durable than 6-3 at low IPR** because all blades share load
- A 6-blade F-Type is approximately equivalent to running a **5-bladed bit**
- Radial shift variants exist to make F-Type more durable (infinite variants possible)
- C2 F-type rippers engage more material than C1 F-type → C2 F-Type is more aggressive than C1

### Dragon Layout
- Trailing backups have their **own dedicated blades**
- Prevents pack-off → maintains ROP
- Backups don't share blade with primaries so cuttings evacuate freely

### Mixed Cutlet Type
- Employs **all three cutlet shapes** (Forward, Reverse, Redundant) in one layout
- Aggressive, similar in nature to F-type aggressiveness
- May be more efficient at dislodging/fracturing rock
- Forces act in **multiple directions simultaneously** with different magnitudes

---

## 4. Ripper / Cutter Configurations

| Config | Description | Aggressiveness |
|--------|-------------|---------------|
| **C0** | Full round PDCs in all positions | Baseline (least aggressive) |
| **C1** | Ripper cutters on secondary blades only (no rippers in cone) | More aggressive than C0 |
| **C2** | Alternating ripper cutters on primary blades | More aggressive than C0 |

- All ripper configurations (C1, C2, mixed) are **more aggressive than full round cutters** in every position
- Whether C1 or C2 has more rippers is **case-by-case** depending on the specific bit design

---

## 5. Backup Element Placement & Engagement

### Why Backup Elements Exist (3 purposes)

1. **Prevent DBR**: Underexposed trailing backups pack off when primaries wear down. When primaries fail past a threshold, secondaries begin working → torque increases, ROP decreases → signals driller to pull bit before DBR occurs.

2. **Limit ROP**: Improve steerability when sliding; reduce straight-hole deviation; prevent primary cutter overload (crushing/shearing).

3. **Increase Diamond Density**: Working/active backups increase durability at the cost of ROP.

### Placement Components & Their Effects

| Component | Definition | Decrease → Effect | Increase → Effect |
|-----------|-----------|-------------------|-------------------|
| **Radial Tip Offset** | Distance between largest radial point on primary vs secondary | Min engage speed ↓, Magnitude ↑, Rate ↓ | Min engage speed ↑, Magnitude ↓, Rate ↑ |
| **Z Tip Offset** | Axial distance between tip heights of primary and secondary | Min engage speed ↓, Magnitude ↑ | Min engage speed ↑, Magnitude ↓ |
| **Degrees Trailing** | Angular offset between primary and secondary elements | Less sensitive to IPR changes | Min engage speed ↓, Magnitude ↑, Rate ↑ (more sensitive to IPR changes) |
| **Element Diameter** | Size of secondary relative to primary | N/A | Smaller secondary → Magnitude ↓, Rate ↓ |
| **Element Vectors** | Combination of tilt and backrake | Shifts cutlet centroid location in combination with radial offset and diameters |

### When Secondary Elements Engage (constant input forces assumed)

- Working element count increases → **increased bit durability**
- Forces on primaries decrease → **likely increases footage** (depends on cutter failure mechanism)
- Reduced primary overload risk → **fewer crushed elements**
- Instantaneous ROP decreases → **does NOT necessarily reduce overall ROP**
- Input forces (WOB, torque) are **redistributed among more cutlets**

### When Knuckles Engage

- **Limit ROP** (primary effect)
- Improve steerability when sliding
- Reduce straight-hole deviation
- Prevent primary cutter overload
- Trade-off: ROP reduction vs. durability and directional control

### Engagement Progression with IPR

- At **low IPR**: Only the most exposed cutters work → aggressive behavior, fewer cutters sharing load
- At **high IPR**: Backup cutters engage → load redistribution → more durable, lower instantaneous ROP
- The engagement_spread metric captures the IPR range over which cutters progressively engage
- A wider spread = more gradual load redistribution = better durability progression

---

## 6. PDC-Rock Interaction Fundamentals

### Depth of Cut (DOC) Progression
1. **Chamfer engages first** at lower DOC
2. **Cutter face engages** as DOC increases

### Chamfer vs Face Performance
- **Face**: More efficient at shearing rock
- **Chamfer**: Increases compression on diamond → increases tangential force (torque) required to exceed element yield strength
- **Balance**: Delicate balance between chamfer and face engagement; affected by drilling parameters, BHA design, PDC properties, and many other variables

### Effective Rake Angle
- As **helix angle increases**, effective rake angle (relative to POM) **decreases**
- Helix angle increases with higher IPR
- Helix angle decreases at larger radii
- Therefore: **gauge cutters have higher effective rake** at a given IPR than inner cutters → gauge cutters are more aggressive relative to inner cutters

---

## 7. Helical Path of Motion (POM) & IPR

### POM Definition
- During drilling, every cutter traces a **helical path**
- Two variables define the helix:
  1. **Radius**: As radius increases → helix angle decreases
  2. **IPR**: As IPR increases → helix angle increases

### IPR Formulas
```
IPR (in/rev) = ft/hr × 12 / (60 × RPM)
IPR (in/rev) = ft/hr × RPM × 0.2   [simplified variant used in documents]
ft/hr at 100 RPM ≈ IPR × 500
```

### Key IPR Relationships
- Higher IPR → more cutters engage → load distribution shifts
- Higher IPR → larger helix angle → lower effective rake angle
- At larger radii, helix angle is naturally lower → higher effective rake → more aggressive gauge behavior

---

## 8. Cutlet Plots & Force Calculations

### What Cutlet Plots Show
- 2D revolved projection simulating drilling at a given IPR
- Shows the **portion of each element doing work** at that IPR
- Cutlet area and centroid determine rock interaction volume
- Element orientation relative to POM determines force direction

### Force Components
Applying tangential psi (rock yield strength) to each cutlet calculates:
- **Tangential forces** → torque
- **Axial forces** → WOB
- **Radial forces** → lateral forces on bit body

Sum of forces predicts: WOB, Torque, magnitude/direction of lateral force

### Limitations
- 2D simplification of 3D process
- Assumes constant rock properties and uniform stress distribution
- In practice these conditions vary significantly

---

## 9. Durability Scoring Model

**Scale**: 0 = most aggressive, 9 = most durable

### Fundamental Principle: Constant Volume

For a given hole size (bit diameter) and IPR, the **total volume of rock removed per revolution is constant** (π × r² × IPR). This volume is divided among all working cutlets. Therefore:
- **More cutlets = smaller individual cutlets = less force per cutter = more durable**
- **Fewer cutlets = larger individual cutlets = more force per cutter = more aggressive**

This is the single most important physical relationship for durability scoring.

### Primary vs Secondary Blade Analysis

In a typical 6-bladed design:
- **Primary blades** (typically 1, 3, 5): Start near the center of the bit, cover the full radial profile
- **Secondary blades** (typically 2, 4, 6): Start closer to the middle of the bit's radius, do NOT extend to center

At radial positions where BOTH primary and secondary blades have row 1 (primary) cutters, the **cutlet size differential** is a key durability indicator:
- If primary blade row 1 cutlets are much LARGER than secondary blade row 1 cutlets at overlapping radii → primary blades doing most of the work → more aggressive → less durable
- If they are similar size → work is well distributed → more durable

### Working Row 2 (Secondary/Backup) Cutters

Row 2 cutters that actually form cutlets (i.e., they are doing work at operating IPR) **increase durability** because:
- More cutters sharing the same constant total volume of rock
- Forces on primaries decrease
- Reduced primary overload risk
- They also **decrease cutting efficiency** and **decrease aggressiveness**

### Components & Weights (Cutlet-Derived, Global Scoring)

| Component | Weight | What It Measures | Higher = More Durable |
|-----------|--------|-----------------|----------------------|
| **Cutter Density** | 22% | Cutlets per sq inch of bit face area | More cutlets sharing constant volume = less force each |
| **Load Balance** | 18% | 1 - Gini coefficient of cutlet areas | Even per-cutter load = no single cutter overloaded |
| **Chamfer Toughness** | 15% | Average chamfer fraction (perimeter×0.020"/area) | Smaller cutlets have proportionally more chamfer = tougher |
| **Peak Moderation** | 12% | 1/max_mean_ratio (largest vs average cutlet) | No single cutter doing disproportionate work |
| **Blade Balance** | 10% | 1/(1+blade_cv) of total work per blade | Even per-blade load distribution |
| **Backrake** | 10% | Average backrake / 30° (capped at 1.0) | Higher backrake = more conservative cutting angle |
| **Backup Engagement** | 7% | Fraction of cutlets from row 2+ cutters | Working backups sharing load at operating IPR |
| **Pri-Sec Balance** | 6% | 1/pri_sec_area_ratio at overlapping radii | Balanced primary-secondary blade work distribution |

### Scoring Method: Global Min-Max Normalization

All metrics are size-independent (ratios, density per unit area, CV, entropy) so they are directly comparable across all bit sizes. Scoring uses global min-max normalization across the entire population — no per-size-bucket rescaling (which amplified small differences within small groups).

---

## 10. Steerability Scoring Model

**Scale**: 0 = least steerable, 9 = most steerable

### Components & Weights

| Component | Weight | What It Measures | Higher = More Steerable |
|-----------|--------|-----------------|------------------------|
| **Axial Dominance** | 16% | Ratio of axial force components | Axial forces = better tool face control |
| **Gauge Openness** | 12% | Inverse of gauge cutter ratio | Fewer gauge cutters = less resistance to side forces |
| **Gauge Aggressiveness** | 12% | Inverse of gauge backrake | Lower gauge backrake = less stabilizing |
| **Cone Aggressiveness** | 12% | Inverse of cone backrake | Lower cone backrake = more aggressive cone = builds angle |
| **Knuckle Effect** | 12% | Effective knuckle ratio (adjusted for non-working) | Knuckles limit ROP → improve steerability when sliding |
| **Profile Depth** | 8% | Z range of cutter positions | Larger Z range = more profile = more steerable |
| **Blade Factor** | 8% | Inverse of effective blade count | Fewer blades = less stabilizing |
| **Siderake** | 8% | Average absolute siderake | More siderake = directional force component |
| **Gauge Backup Openness** | 7% | Inverse of gauge backup ratio | Fewer gauge backups = less stabilizing at gauge |
| **Lateral Force** | 5% | Lateral force imbalance | Higher imbalance = tendency to walk |

### Key Steerability Relationships
- **Reverse Spiral** layouts have predominantly axial forces → best tool face control → most steerable
- **Fewer gauge cutters** = gauge region is more "open" = less stabilizing = easier to steer
- **Knuckles** limit ROP which **improves slide steerability** (lower ROP = motor has more control)
- **Non-working knuckles** (extremely underexposed, Z offset > 3x average) don't contribute to steerability

---

## 11. Key Metrics Definitions

### Layout Geometry Metrics

| Metric | Range | Definition |
|--------|-------|-----------|
| `redundancy_score` | 0–1 | Average of top N/2 blade-pair radial overlap ratios (Jaccard on binned radial positions) |
| `axial_force_ratio` | 0–1 | Mean |dz|/magnitude across all cutter orientation vectors |
| `radial_force_ratio` | 0–1 | Mean sqrt(dx²+dy²)/magnitude across all orientation vectors |
| `perpendicular_force_ratio` | 0–1 | 1.0 - |axial - radial| (balanced = perpendicular to profile) |
| `six_three_offset` | 0+ inches | Mean Z difference between primary and secondary blade groups at nose |
| `blade_exposure_equality` | 0–1 | How uniformly blades cover the radial profile (accounts for intentional short secondary blades in redundant designs) |
| `effective_blade_count` | integer | Number of blades starting from near center (<15% gauge radius) |

### Cutter Geometry Metrics

| Metric | Definition |
|--------|-----------|
| `total_cutters` | Total count of unique cutters in ME file |
| `num_blades` | Number of distinct blades |
| `avg_backrake` | Mean backrake across all cutters (degrees) |
| `avg_cone_backrake` | Mean backrake in cone zone (<30% gauge radius) |
| `avg_nose_backrake` | Mean backrake in nose zone (30-60% gauge radius) |
| `avg_gauge_backrake` | Mean backrake in gauge zone (>85% gauge radius) |
| `cutter_density` | Cutters per unit radius, size-adjusted (16mm reference) |
| `avg_spacing_cv` | Mean coefficient of variation of per-blade cutter spacings |
| `gauge_cutter_ratio` | Fraction of cutters in gauge zone (>85% radius) |
| `avg_siderake` | Mean absolute siderake (degrees) |
| `lateral_resultant` | Magnitude of net lateral force vector / cutter count |
| `z_range` | Max Z - Min Z across all cutters (profile depth in inches) |
| `avg_primary_cutter_dia_mm` | Mean diameter of primary (row 1) cutters in mm |

### Backup & Engagement Metrics

| Metric | Definition |
|--------|-----------|
| `backup_ratio` | n_backup / total_cutters |
| `knuckle_ratio` | n_knuckle_backups / n_backup |
| `n_primary` | Count of row-1 (primary) cutters |
| `n_backup` | Count of row-2+ (backup) cutters |
| `n_knuckle_backups` | Count of CPS/knuckle-type backups |
| `n_pdc_backups` | Count of PDC-type backups |
| `avg_radial_offset` | Mean radial distance between paired primary-backup cutters |
| `avg_z_offset` | Mean axial (Z) distance between paired primary-backup cutters |
| `max_z_offset` | Maximum Z offset among all backup pairs |
| `paired_positions` | Number of backup cutters successfully paired with a primary |
| `backup_profile_coverage` | Fraction of primary radial bins that also have a backup |
| `engagement_spread` | IPR range over which cutter classes progressively engage |
| `median_threshold` | Median engagement threshold (in/rev) from ME Settings |
| `low_ipr_fraction` | Fraction of thresholds at or below median |
| `non_working_ratio` | Fraction of backups with Z offset > 3x average (essentially decorative) |
| `gauge_backup_ratio` | Fraction of backups in gauge zone (>85% radius) |
| `avg_primary_backrake` | Mean backrake of primary cutters only |
| `avg_backup_backrake` | Mean backrake of PDC backup cutters only |
| `backrake_differential` | avg_backup_backrake - avg_primary_backrake |

---

## 12. Design Trade-offs & Rules of Thumb

### Aggressiveness vs Durability (Inverse Relationship)
- **More aggressive** = lower backrake, fewer blades, less redundancy, larger cutters, more open gauge
- **More durable** = higher backrake, more blades, high redundancy, smaller cutters, more backups

### Specific Trade-offs

1. **6-3 vs F-Type**: Both similarly aggressive, but F-Type is more durable at low IPR because all blades are equally exposed. 6-3 is vulnerable at low IPR (acts as 3-blade bit).

2. **PDC Backups vs Knuckle Backups**:
   - PDC backups: Increase diamond density → more durable → reduce ROP proportionally
   - Knuckles: Limit ROP more aggressively → improve steerability when sliding → prevent primary overload

3. **Cutter Size**:
   - Smaller cutters (11-13mm): More durable (less exposure per cutter, more cutters fit), but lower ROP potential
   - Larger cutters (16-19mm): More aggressive (more exposure), higher ROP potential, but less durable

4. **Backrake by Zone**:
   - Low cone backrake: Builds angle (steerable)
   - Low gauge backrake: Less stabilizing (steerable but less durable at gauge)
   - High gauge backrake: More stabilizing (durable at gauge but less steerable)

5. **Backup Backrake Differential**:
   - Higher differential (backup BR >> primary BR): Backups engage more conservatively; gradual load transition
   - Lower differential: Backups engage similarly to primaries; sharper transition

6. **Instantaneous ROP vs Overall ROP**:
   - When backups engage, instantaneous ROP decreases
   - This does NOT necessarily reduce overall ROP (bit lasts longer, fewer trips)

7. **Blade Count**:
   - Fewer blades = more aggressive, more steerable, less stable
   - More blades = more durable, more stable, harder to steer
   - F-Type 6-blade ≈ effective 5-blade aggressiveness

8. **Forward vs Reverse Spiral**:
   - Forward: Radial forces → more torque → surface bits, limited WOB scenarios
   - Reverse: Axial forces → more WOB → curves/laterals, better tool face control

---

## 13. Min Engagement File Structure

### Assy.Model Sheet (Cutter Data)
- **Row 1, Col C**: Bit diameter (e.g., "8.75 inch")
- **Rows 9+**: One row per cutter entry (tip and base; deduplicated by keeping first occurrence per name)

| Column | Content |
|--------|---------|
| B | X position |
| C | Y position |
| D | Z position |
| E | dx (orientation vector) |
| F | dy (orientation vector) |
| G | dz (orientation vector) |
| H | Pocket radius |
| I | Pocket depth |
| J | Cutter name (e.g., "1.103") |
| K | Element type (PDC or CPS/knuckle) |
| L | Zone |
| Q | Theta (degrees) |
| R | Radial position |
| AN | Tilt (raw) |
| AO | Backrake (raw; v6 files store as 180 - actual) |
| AP | Siderake (raw; v6 files may need abs(180 - raw) conversion) |

### Assy.Model Summary Area (Rows 29-30)
Pre-computed engagement values for secondary cutting structure:
- **Row 29**: Sec PDC engagement (nose and taper)
- **Row 30**: Knuckle engagement (nose and taper)
- **v6.xx**: Columns JH (Nose), JI (Taper), label in JA
- **v7.xx/v8.xx**: Columns JJ (Nose), JK (Taper), label in JC

### Settings Sheet
- **Row 19, Col C**: Bit diameter (fallback)
- **Rows 14+**: Engagement thresholds per cutter class
  - Column M: Min in/rev for that class to engage
  - Column P: Threshold in/rev (preferred over M if available)

### Version-Specific Conversions
- **v6.xx backrake**: Stored as `180 - actual_backrake` → convert: if raw > 90, backrake = 180 - raw
- **v6.xx siderake**: May need `abs(180 - raw)` if raw > 90
- **v7.xx / v8.xx**: Values stored directly (no conversion needed)

---

## 14. Cutter Naming Convention

Format: **X.YZZ**
- **X** = Blade number (1-9+)
- **Y** = Row number:
  - 1 = Primary cutter
  - 2+ = Backup/trailing cutter
- **ZZ** = Radial position number along the blade

Examples:
- `1.103` = Blade 1, Row 1 (primary), Position 03
- `3.205` = Blade 3, Row 2 (backup), Position 05

Entries to skip:
- Names containing "flip", "vector", or "Part"
- Names without a "." separator

---

## 15. Data Sources & File Inventory

### Critical Files (Ranked by Importance)

1. **compute_ratings.py** (1,436 lines) — All rating algorithms and metric definitions
2. **bit_ratings_analysis.json** — Scored outputs for 130 bit designs with 45+ raw metrics each
3. **Bit Designs/*/\*Min Engagement\*.xlsm** (188 files) — Source data for all geometry metrics
4. **Design Types/General Characterization of Design Layouts and Ripper Configurations.docx** — Layout type definitions and force characteristics
5. **Design Types/Characterization of trailing back up cutter placement.docx** — Backup element theory, engagement mechanics, placement effects
6. **Bit Applications Tool r2.xlsx** — Main workbook with bit identification, ratings (AO-AP), engagement values (BA-BE)

### Repository Statistics

| File Type | Count | Location |
|-----------|-------|----------|
| XLSM (Min Engagement) | 188 | Bit Designs + Design Types |
| XLSX (analysis) | 108 | Bit Designs |
| DOCX (design docs) | 217 | Bit Designs + Design Types |
| PDF (spec sheets) | 486 | Bit Designs + Original Assy Spec Sheets |
| Python scripts | 1 | Root |
| JSON (metrics) | 1 | Root |
| **Total** | **815** | |

### ME File Version History
- **v6.02–v6.08**: Older format; backrake stored as 180-actual; different summary column positions (JH/JI)
- **v7.01**: Intermediate format; direct backrake values; summary columns JJ/JK
- **v8.01**: Latest format; same conventions as v7

---

## Appendix: Scoring Normalization Process

Both durability and steerability scores use the same normalization pipeline:

1. **Compute raw component values** for each bit from ME file geometry
2. **Normalize each component to 0–1** across the full population (min-max scaling)
3. **Weighted sum** of normalized components produces a raw composite score
4. **Scale to 0–9** using min-max across the population

This means scores are **relative** — a bit scored 9.0 is the most durable *in this population*, not an absolute measure. Adding new bits to the population may shift all scores.
