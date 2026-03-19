# Cutlet Analysis & Force Calculation — Complete Method Notes

## Reference for replicating the ME file's cutlet-based force analysis computationally

> **Source**: Reverse-engineered from `2176r06 Min Engagement v8.01.xlsm`
> (pushed to main as the example/reference file)

---

## 1. THE BIG PICTURE

The Min Engagement file does three things:

1. **Imports raw bit geometry** from Solid Edge (columns B-J)
2. **Simulates drilling** at a given IPR to produce cutlet shapes (columns AS-CF)
3. **Calculates forces** from those cutlet shapes (columns CG-IE)
4. **Determines backup engagement** — at what drilling speed backups begin working (columns IG-JK)

The **critical bottleneck** is step 2: computing the cutlet geometry (centroid + area).
Currently this requires either AutoCAD (MASSPROP command on drawn cutlet shapes)
or Solid Edge (3D model viewed orthogonally). We need to replicate this computationally.

---

## 2. COLUMN MAP — Assy.Model Sheet

### 2.1 Raw Solid Edge Import (B-J)

| Column | Content | Notes |
|--------|---------|-------|
| B | X position (absolute) | 3D Cartesian, World CS |
| C | Y position (absolute) | |
| D | Z position (absolute) | |
| E | Delta X (orientation vector) | Unit direction of cutter axis |
| F | Delta Y | |
| G | Delta Z | |
| H | Pocket radius | Three sizes observed: 0.3125, 0.2645, 0.255 |
| I | Pocket depth | Either 0.52 or 0.415 |
| J | Cutter name | Format: blade.row_position (e.g., 1.101, 2.229) |

- **Row 1, Col C**: Bit diameter and design ID (e.g., "7.875 2176r06")
- **Rows 1-8**: Headers
- **Rows 9+**: One entry per cutter (tip then base; ~42 entries for 2176r06)

### 2.2 Cutter Positions & Orientations (M-AQ)

| Column | Content | Notes |
|--------|---------|-------|
| M | CS-Name | Cutter station ID (same as J but in computed section) |
| N-P | X, Y, Z | Cartesian position relative to World CS |
| Q | Theta (degrees) | Angular position in polar coords |
| R | r (radial distance) | Distance from bit center |
| S | Z | Same Z as column D |
| T-V | X, Y, Z | Cartesian relative to reference cutter 1.112's theta |
| W-Y | Theta, r, Z | Polar relative to 1.112's theta |
| Z-AB | Delta X, Y, Z | Orientation vector (Cartesian, World CS) |
| AC | Vector length | Always 1.0 (unit vector) |
| AD-AF | Delta X, Y, Z | Unit vector (normalized) |
| AG-AH | Rotation X°, Y° | Successive rotations (X then Y axis) |
| AI-AJ | Machine angles A°, C° | Vari-Axis 630 machine tool |
| AK-AM | Delta X, Y, Z | Polar Z Projection to XZ Plane unit vector |
| **AN** | **Tilt (degrees)** | X rotation angle |
| **AO** | **Back Rake (degrees)** | Y rotation angle |
| **AP** | **Side Rake (degrees)** | Projection angle |
| **AQ** | **Helix Angle (degrees)** | At this cutter's radial position |

### 2.3 Drilling Speed Inputs (AS-AT)

| Row | AS | AT | Meaning |
|-----|----|----|---------|
| 3 | 100 | rpm | Reference RPM |
| 4 | 125 | ft/hr | Reference drilling speed |
| 5 | 0.25 | in/rev | **IPR = ft/hr × 12 / (60 × RPM)** |
| 6 | 0.000694 | in/deg | **IPR / 360** |

For data rows (9+):
- **AS** = Offset Angle (degrees) — angular position of this cutter
- **AT** = Cutlet Z Offset (inches) — helical Z displacement at this angle

### 2.4 THE CRITICAL CUTLET INPUTS (AV-AY)

| Column | Content | Source |
|--------|---------|--------|
| AV | CS-Name | Cutter identifier |
| **AW** | **Centroid Y (Radial)** | From AutoCAD MASSPROP or Solid Edge |
| **AX** | **Centroid Z** | From AutoCAD MASSPROP or Solid Edge |
| **AY** | **Cutlet Area (in²)** | From AutoCAD MASSPROP or Solid Edge |

**These three values (AW, AX, AY) are THE bottleneck.** They currently require
CAD software to compute. Everything downstream is pure math from these inputs.

**Key observation**: Only primary cutters (row 1, e.g., 1.101, 3.103) have cutlet data.
Backup cutters (row 2+, e.g., 2.229) do NOT have cutlet data at 0.25 IPR because
they haven't engaged yet. Rows 15-19, 23-26, 32-36, 40-43, 50 are empty in AV-AY.

### 2.5 Derived Geometry (AZ-CF)

| Column | Content |
|--------|---------|
| AZ | Volume (in³) — computed from Area × radius × circumference |
| BA | % of total volume |
| BB | Z offset: PDC center → cutlet centroid |
| BC-BD | Polar coords of centroid (theta°, r) |
| BE | Centroid Z |
| BF-BH | PDC face plane: Point 1 (X, Y, Z) |
| BI-BK | PDC face plane: Point 2 (P1 + v) |
| BL-BN | PDC face plane: Point 3 (P1 + w) |
| BO-BR | PDC face plane equation: Ax + By + Cz + D = 0 |
| BS-BU | Centroid Z plane: Point 1 |
| BV-BX | Centroid Z plane: Point 2 |
| BY-CA | Centroid Z plane: Point 3 |
| CB-CE | Centroid Z plane equation (always A=0, B=0, C=-1, D=Z_centroid) |
| CF+ | Direction vector of PDC plane ∩ Centroid Z plane intersection line |

### 2.6 3D Intersection Geometry (CG-ER)

**Columns CG-DL: Line/Sphere and Line/Cylinder intersections**
- CG-CN: Two points defining a line (P1, P2)
- CO: Sphere radius (PDC radius)
- CP-CS: Quadratic coefficients (a, b, c, d) for sphere intersection
- CT-CU: Parameter solutions (U1, U2)
- CV-DA: Intersection points IP1 and IP2 (segment endpoints on PDC face)
- DB-DG: Cylinder definition (always centered on Z-axis: CP1=(0,0,-10), CP2=(0,0,10))
- DH: Cylinder radius (= centroid radial position)
- DI-DK: Cylinder vector (always (0,0,20))

**Columns DM-ER: Cylinder intersection solve + final centroid**
- DM-DW: Dot products (md, nd, dd, nn, mn, mm)
- DX: k scalar
- DY-EB: Quadratic coefficients for cylinder intersection
- EC-ED: Parameter solutions (t1, t2)
- EE-EJ: Intersection points at t1 and t2
- EK-EL: Boolean — is intersection on the segment? (t2 is always True)
- **EM-EO**: Final 3D centroid in Cartesian (X, Y, Z)
- **EP**: Centroid theta (degrees)
- **EQ**: Centroid r (radial — should match AW)
- **ER**: Centroid Z (should match AX)

### 2.7 Vector Force Calculations (ES-FX)

| Columns | Content |
|---------|---------|
| ES-EU | Cutlet Face unit vector (Cartesian, World CS) |
| EV | Polar Z Projection rotation angle |
| EW-EY | Rotated unit vector |
| EZ | Helix angle at this cutter's position |
| FA-FC | Path of Motion (POM) unit vector |
| FD-FF | POM Polar Z Projection |
| FG-FI | r vector = Face unit vector - POM unit vector |
| **FJ** | **Rake Total (degrees)** — angle between face and POM |
| FK-FM | Projected vector lengths in XY plane |
| **FN** | **Rake Perpendicular to Z (degrees)** |
| FO | Negative rake flag (+1 or -1) |
| **FP** | **Side Rake Perpendicular to Z** |
| FQ-FT | Force component percentages (Perp→Z and Torque, X and Y) |
| FU-FW | Projected vector lengths in XZ plane |
| FX | Rake in XZ plane |

### 2.8 Final Force Computation (FY-HJ)

| Columns | Content |
|---------|---------|
| FY | Negative rake flag |
| FZ | Helix rake in XZ plane |
| GA | 2nd rotation correction |
| GB-GD | Force direction components (Perp→Z, Tangential→Z, Axial) |
| GF-GI | Percentage of shear force per direction |
| GJ | Cutlet centroid radius (in) |
| GK | Helical circumference (ft) |
| GL | Cutlet area (in²) |
| GM | Volume (in³) |
| GN | % of total volume |
| GO-GQ | Side rake, helix rake, total rake (repeated for this section) |
| GR | Per-cutter sum of force components (ft·lbs) |
| GS-GZ | Radial/Tangential/Axial forces in ft·lbs and lbs |
| **GW** | **29,500 psi — formation yield strength (tangential force)** |
| **GX** | **Torque (ft·lbs)** |
| HA-HD | Lateral force: lbs, vector angle, X component, Y component |
| HE-HH | Tangential force: lbs, vector angle, X/Y components |
| HI | CS-Name |

### 2.9 Force Results Summary (HK-IE)

**Row 6 contains totals. Row 7-8 are headers. Rows 9+ are per-cutter data.**

| Column | Content | Units |
|--------|---------|-------|
| HK | Radial position | ordinal |
| HL | Element name | e.g., 1.101 |
| HM | Side Rake (XY Proj.) | degrees @ 0 IPR |
| HN | Back Rake (Y→X Suc. Rot.) | degrees @ 0 IPR |
| HO | YZ Projection | degrees |
| HP | (angle @ IDS) | degrees |
| HQ | Face↔POM | degrees @ 0 IPR |
| HR | (Face↔POM @ IDS) | degrees |
| **HS** | **Centroid Radius** | in |
| **HT** | **Cutlet Area** | in² |
| **HU** | **Volume at IDS** | in³ |
| HV | % of total volume | % |
| **HW** | **Yield Strength** | psi (always 29,500) |
| **HX** | **Torque** | ft·lbs |
| HY | Radial force | ft·lbs (X component) |
| **HZ** | **Tangential force** | ft·lbs (Y component) |
| **IA** | **Axial force** | ft·lbs (Z component) |
| IB | Σ Cutlet forces | ft·lbs (sum of components) |
| IC | Radial force | lbs (X component) |
| **ID** | **Tangential force** | lbs (Y component) |
| **IE** | **Axial force** | lbs (Z component) |

**Row 6 totals for 2176r06 @ 0.25 IPR:**
- Torque: **4,732 ft·lbs** (customer reports ~8,500 at the motor)
- Tangential work: **29,746 ft·lbs**
- Axial force: **6,085 lbs**
- Tangential force: **28,647 lbs**
- Lateral force: **1,926 lbs @ 95.6° relative to Blade 1**

### 2.10 Backup Engagement Calculation (IG-JK)

| Column | Content |
|--------|---------|
| IG | Nose flag (0.75 threshold) |
| IH | Gage diameter flag (-0.0625) |
| II | Cutter name |
| IJ | Hide flag |
| IK | "Bu?" — is this a backup element? |
| IL | Primary cutter (paired) |
| IM | Secondary cutter (this backup) |
| IN | Degrees trailing |
| IO | Exposure |
| IP | Engagement threshold (in/rev) |
| **IQ** | **Engagement speed (ft/hr per 100 rpm)** |
| IR-IV | Primary cutter ellipse: Center X/Y, Major r, Minor r, Tilt |
| IW-JA | Secondary cutter ellipse: Center X/Y, Major r, Minor r, Tilt |
| JC | Element type (part number or "CPS-xxxxx") |
| JD | Sigma count |

**Summary statistics (rows 20-24, 29-30, columns JC-JK):**

| Row | Type | Nose (JJ) | Taper (JK) |
|-----|------|-----------|------------|
| 29 | Sec PDC (1313) | **430 ft/hr** | **470 ft/hr** |
| 30 | Knuckle (CPS-32193r1) | — | **640 ft/hr** |

---

## 3. OTHER SHEETS

### 3.1 Settings Sheet

- **Rows 1-6**: Blade colors (9 blades supported, RGB strings)
- **Row 9**: C=9 (data starts at row 9), plus RPM table (50-250 in steps)
- **Rows 13-20**: IPR/feed rate table — 8 threshold classes:
  - 0.0625, 0.09375, 0.140625, 0.210938, 0.316406, 0.474609, 0.711914 in/rev
  - Each maps to a movement distance for AutoCAD script generation
- **Row 19**: C=7.875 (bit diameter)
- **Rows 14-27**: Part catalog — cutter specifications:
  - Part 1108: 0.433" dia, 0.315" length
  - Part 1213: 0.481" dia, 0.52" length
  - Part 1313: 0.529" dia, 0.52" length (used in 2176r06)
  - Part 1613: 0.625" dia, 0.52" length
  - Through Part 1913: 0.75" dia
- **Rows 28-36**: Special elements — TCI domes and PDC domes:
  - CPS-32193r1: 0.510" TCI Dome (active knuckle in 2176r06)

### 3.2 MassProp Sheet

**This is where AutoCAD MASSPROP extraction data lands.**

| Column | Content |
|--------|---------|
| A | Pattern index (from formula) |
| B | Cutter name (e.g., 1.101) |
| C | Centroid X |
| D | Centroid Y |
| E | Area |
| H-L | Cleaned/sorted mirror of A-E |

26 entries matching the 26 primary cutters. These values feed AW-AY.

### 3.3 Settings (2) Sheet

AutoCAD script generation configuration:
- Column mappings for which Assy.Model columns to read
- Feed rate / IPR table with RPM-to-speed calculations
- Draw flags (checkmarks) for which engagement levels to plot
- Gage diameter: 6.375"

### 3.4 Sheet1

Cutter-to-position lookup table:
- Maps cutter names to sequential IDs
- Identifies primary (row 1) vs backup (row 2) cutters
- Positions 01-10: primary only; Positions 11+: both primary and backup

---

## 4. THE FORCE CALCULATION METHOD

### 4.1 Core Formula

For each cutter at a given IPR:

```
Force = Cutlet_Area × Formation_Yield_Strength
```

Where Formation Yield Strength = 29,500 psi (tangential).

This total force is then decomposed into components using the cutter's
orientation vectors relative to the Path of Motion (POM):

```
Tangential_Force = Force × sin(rake_angle) × direction_factors
Axial_Force      = Force × cos(rake_angle) × direction_factors
Radial_Force     = Force × side_rake_factors
Torque           = Tangential_Force × centroid_radius
```

### 4.2 Volume Calculation

```
Volume = Cutlet_Area × 2π × Centroid_Radius (per revolution)
       = Area × helical_circumference_at_radius
```

### 4.3 Rake Angle Decomposition

The file computes multiple rake angle representations:
1. **Total Rake** (FJ): Angle between cutter face and POM — the "true" effective rake
2. **Rake Perp→Z** (FN): Component perpendicular to the Z-axis
3. **Side Rake Perp→Z** (FP): Lateral component
4. **XZ Plane Rake** (FX): Projection onto the axial plane

The force split uses these angles to decompose total force into:
- **GB**: Perpendicular to Z (radial force on bit)
- **GC**: Tangential to Z (torque-producing)
- **GD**: Axial (WOB-producing)

### 4.4 Path of Motion (POM) Unit Vector

At each cutter's radial position:
```
helix_angle = atan(IPR / (2π × radius))
POM_vector = [tangential_component, 0, axial_component]
           = [cos(helix_angle), 0, -sin(helix_angle)]  (in local frame)
```

The POM is then rotated into the World CS using the cutter's angular position.

---

## 5. THE CUTLET GEOMETRY CHALLENGE

### 5.1 What a Cutlet Is

A **cutlet** is the cross-sectional shape of rock that a single cutter removes
at a given IPR. It's the intersection of:

1. The **PDC cutter face** (a circle tilted in 3D by backrake and siderake)
2. The **rock surface** at this cutter's radial position after preceding cutters
   on other blades have already cut

### 5.2 How It's Currently Computed

**Method 1 — AutoCAD (no longer available):**
1. An Excel macro generates an AutoCAD script file
2. The script draws the "Revolved Z Projection" — all cutters projected onto
   a single radial-Z plane, with helical offsets applied at the given IPR
3. Each cutter appears as an ellipse (circle tilted by rake angles)
4. The portion of each ellipse that extends beyond preceding cutters = the cutlet
5. AutoCAD MASSPROP command extracts centroid and area for each cutlet
6. Another macro reads the MASSPROP log file and places values into Excel (MassProp sheet)
7. Values are then copied to columns AW-AY

**Method 2 — Solid Edge (current):**
1. A macro places 3D cutter bodies into Solid Edge
2. Right-side orthogonal view gives an "equivalent" abstract cutlet plot
3. Visual inspection only — doesn't directly quantify centroid/area

### 5.3 What We Need to Replicate Computationally

For each primary cutter, compute:
- **Centroid (radial, Z)**: The geometric center of the cutlet shape
- **Area (in²)**: The cross-sectional area of the cutlet shape

The cutlet shape is determined by:
1. **Cutter geometry**: PDC diameter (from pocket radius H), backrake (AO), siderake (AP), tilt (AN)
2. **Cutter position**: Radial position (R), angular position (Q), Z position (S)
3. **IPR**: Determines the helix pitch, which determines DOC at each radius
4. **Preceding cutters**: At similar radial positions on earlier blades, their profiles
   define the "already-cut" surface. This cutter only cuts what's left.

### 5.4 The Revolved Z Projection Approach

All cutters are "revolved" onto a single radial-Z plane:
1. Each cutter's angular position (theta) determines its Z offset via the helix:
   `Z_offset = theta × IPR / 360` (= theta × in/deg)
2. The cutter face projects as an ellipse on the radial-Z plane
3. Ellipse parameters depend on PDC diameter, backrake, siderake, and tilt
4. The cutlet for each cutter = the portion of its projected ellipse that
   extends BELOW the profile established by all preceding cutters

This is fundamentally a **2D computational geometry** problem on the revolved plane:
- X-axis = radial position
- Y-axis = Z position (depth)
- Each cutter is an ellipse
- Cutlet = ellipse minus overlap with preceding cutters

### 5.5 Ellipse Parameters for Each Cutter

On the revolved Z projection, each PDC cutter appears as an ellipse:
- **Center**: (radial_position, Z_position - Z_offset_from_helix)
- **Major axis**: PDC diameter × cos(siderake) in radial direction (approximately)
- **Minor axis**: PDC diameter × cos(backrake) in Z direction (approximately)
- **Tilt**: Related to siderake angle

The exact ellipse parameters are more complex (full 3D projection of a tilted circle),
but the ME file's backup engagement section (columns IR-JA) stores primary and
secondary cutter ellipse parameters: Center X/Y, Major r, Minor r, Tilt.

### 5.6 Computing the Cutlet Area and Centroid

For a set of cutters at similar radial positions:
1. Sort by angular position (blade order)
2. First cutter at this radius: cutlet = full projected ellipse clipped to the
   "bottom of hole" profile
3. Each subsequent cutter: cutlet = portion of its ellipse that extends below
   the profile left by the first cutter
4. For each cutlet, compute:
   - Area: ∫∫ dA over the cutlet region
   - Centroid: (∫∫ x dA / Area, ∫∫ z dA / Area)

This can be done with:
- Analytical ellipse-ellipse intersection formulas
- Numerical integration (discretize the ellipse, compute pixel-by-pixel)
- Polygon approximation (approximate ellipses as polygons, use Shapely/similar)

---

## 6. OBSERVATIONS FROM 2176r06 DATA

### 6.1 Cutter Population
- **6 blades** (1-6), though blade 6 only has gauge cutters
- **26 primary cutters** with cutlet data at 0.25 IPR
- **~16 backup cutters** without cutlet data (not engaged at this IPR)
- **3 pocket sizes**: 0.3125" (13mm), 0.2645" (11mm), 0.255" (10mm) radius

### 6.2 Force Distribution
- Inner cutters (cone/nose): Mostly axial force, high Area (~0.06-0.07 in²)
- Mid-radius (nose/taper): Highest volume, balanced forces
- Outer cutters (gauge): Low area (0.002-0.02 in²), mostly tangential
- Cutter 1.101 (center): Area=0.0796 in², Axial=465 lbs
- Cutter 1.107 (r=1.71): Area=0.0697 in², highest volume at 0.7475 in³
- Cutter 5.126 (gauge, r=3.92): Area=0.002 in², only 30 lbs radial

### 6.3 Key Relationships Validated
- Volume increases with radius (longer circumferential path per revolution)
- Area tends to decrease near gauge (less DOC, higher helix angle)
- Torque ≈ Tangential_Force × Radius
- Total predicted torque (4,732 ft·lbs) is ~56% of customer-reported motor torque
  (8,500 ft·lbs) — reasonable given motor losses and BHA friction

---

## 7. VBA MACRO ANALYSIS — CUTLET PLOT METHODS

### 7.1 Overview: Two Methods

The ME file contains two distinct cutlet plot methods, each with its own VBA modules:

1. **AutoCAD Script Method** (`SCR_Output_Module.bas`, `Ellipse.cls`) — generates a `.scr` file
   that draws 2D ellipses in AutoCAD, creates boundary regions, and runs MASSPROP to extract
   centroid and area. Results are parsed from the AutoCAD log file and placed on the MassProp sheet.

2. **Solid Edge 3D Method** (`SE_Cutlets.bas`, `SE_Cutlets_v2.bas`) — places actual 3D cutter
   solid bodies into Solid Edge at coordinate system positions from Excel. Produces a 3D cutlet
   model viewed orthogonally. Does NOT directly compute centroid/area.

**For computational replication, the AutoCAD method is the relevant one** — it's already a 2D
ellipse-based approach that we can replicate with computational geometry.

### 7.2 AutoCAD Script Method — Complete Flow

**Entry Point**: `SCR_Output_Module.ProduceSCRfile()`

Called from the `SCR_file` form with parameters:
- `scrFileLoc`: Output .scr file path
- `Sheet4scr`: Sheet name to read data from (e.g., "Assy.Model")
- `printOption`: "None", "PDF", or "Letter"
- `textHeight`: Text size for labels
- `RunType`: "LOCATION" (static) or "VOLUMETRIC" (with cutlets + MASSPROP)

**Step 1: CollectData()** — Read cutters from Excel into Ellipses[] array

```
Settings sheet defines column mappings:
  Name_Column  → J  (cutter name)
  X_Column     → X  (radial position)
  Z_Column_Stationary → Y  (Z at 0 IPR)
  Z_Column_Drilling   → AT (Z with helical offset at given IPR)
  Rcutter_Column      → H  (cutter diameter)
  DegTilt_Column      → AN (tilt angle)
  DegRake_Column      → AO (back rake angle)
  DegZPolar_Column    → Q  (angular position in degrees)

START_ROW = 9

For each row starting at START_ROW:
  - Read Name, X, Z, Diameter, Tilt, Rake, ZPolar from those columns
  - Create Ellipse object:
      Name = CutterName (as double, e.g., 1.101)
      Xc = -1 × X_Column  (NEGATED — radial position becomes negative)
      Yc = Z_Column_Drilling (for VOLUMETRIC) or Z_Column_Stationary (for LOCATION)
      Major = Rcutter_Column (half-diameter = radius)
      Tilt = -1 × DegTilt_Column (NEGATED)
      Rake = DegRake_Column
      Minor = Major × cos(Rake)  ← KEY: minor axis from back rake
      zPolar = DegZPolar_Column
      RevolutionNum = 1
      MassPropFlag = False
  - Continue until CutterName is no longer numeric
```

**CRITICAL: X is negated and Tilt is negated.** The revolved Z projection has X pointing
left (toward bit center) and the tilt sign is flipped.

**Step 2: AdditionalEllipses(IPR)** — Create revolution 2 and 3 copies (VOLUMETRIC only)

Given N cutters at indices 0..(N-1), where Ellipses(N-1).IsLast = True:

```
Revolution 2 (shifted by 1×IPR, MassProp=True):
  For index = 0 to N-2:
    Ellipses(N + index) = copy of Ellipses(index)
    Shifted down by: Yc = Yc + IPR  (shiftDown(-1 * IPR_Value) adds IPR to Y)
    SolveMassProperties(True) ← THESE are the cutlets that get measured

Revolution 3 (shifted by 2×IPR, MassProp=False):
  For index = 0 to N-2:
    Ellipses(2N-1 + index) = copy of Ellipses(index)
    Shifted down by: Yc = Yc + 2×IPR
    SolveMassProperties(False)
```

**Array layout after AdditionalEllipses:**
- [0..N-1]: Revolution 1 — current positions (N cutters)
- [N..2N-2]: Revolution 2 — shifted by 1×IPR (N-1 cutters, copies of 0..N-2)
- [2N-1..3N-3]: Revolution 3 — shifted by 2×IPR (N-1 cutters, copies of 0..N-2)

Note: The LAST cutter (index N-1) is NOT copied to Rev 2 or Rev 3.

**Physical meaning:**
- Rev 1 = where cutters were on the PREVIOUS revolution (already cut this rock)
- Rev 2 = where cutters are on the CURRENT revolution (these are the cutlets we measure)
- Rev 3 = where cutters will be on the NEXT revolution (provides bottom boundary)

### 7.3 Drawing and Sequencing — THE CRITICAL PART

**Step 3: DrawData() for VOLUMETRIC mode**

```vb
' Find last index
index = 0
Do While Not Ellipses(index).IsLast
    index = index + 1
Loop
TopIndex = index

' === DRAW ELLIPSES IN REVERSE ORDER (last → first) ===
Do While index >= 0
    ' 1) Draw the ellipse
    Call Ellipses(index).Draw(Filenum)

    ' 2) Erase center — trim interior, store boundary points
    Call Ellipses(index).EraseCenter(Filenum, 0.002, 1)

    ' 3) For each LOWER-indexed ellipse, remove this ellipse's
    '    boundary points that fall inside it
    index2 = 0
    Do While index2 < index
        Call Ellipses(index).RemovePointsInEllipse(Ellipses(index2))
        index2 = index2 + 1
    Loop

    index = index - 1
Loop
```

**SEQUENCING RULE: Lower-indexed ellipses MASK higher-indexed ones.**

This means:
- Blade 1 cutters (lowest indices) are DOMINANT — they cut first
- Blade 6 cutters (highest indices) only see what's left after blades 1-5
- Within each blade, inner cutters (earlier indices) have priority over outer cutters
- Rev 1 ellipses (indices 0..N-1) mask EVERYTHING in Rev 2 and Rev 3

The physical interpretation: As the bit rotates, blade 1 encounters the rock first,
then blade 2 (60° later), etc. The helical Z offset in the Y-coordinate accounts
for the bit's descent between blades.

**Step 4: Trim at gauge line**

`TrimTrimmers()` draws vertical lines at X = -GageRadius and X = 0, then trims
everything outside the gauge diameter. The gage line (3.9375 for a 7.875" bit)
defines the outer boundary.

**Step 5: Create boundaries and run MASSPROP**

```vb
' CREATE BOUNDARIES — one per Rev 2 cutter (MassProp=True)
' Process in REVERSE order (same direction as drawing)
Do While index >= 0
    If Ellipses(index).DoISolveMassProperties Then  ' Only Rev 2
        ' Create layer named by cutter number
        createLayer(Ellipses(index).CutterName, "white")
        ' Zoom to the MidThetaPoint of visible arc
        ' Run AutoCAD -BOUNDARY command at the midpoint
        Print #Filenum, "-BOUNDARY"
        Print #Filenum, midThetaPt.x & "," & midThetaPt.y
        ' Turn off the layer (hide the boundary)
    End If
    index = index - 1
Loop

' MASSPROP — turn on each cutter's boundary one at a time
' Process in REVERSE order
Do While index >= 0
    If Ellipses(index).DoISolveMassProperties Then
        ' Turn on layer, zoom, run MASSPROP
        Print #Filenum, "MASSPROP WINDOW ..."
        Print #Filenum, "N"  ' Don't save to file
        ' Turn off layer
    End If
    index = index - 1
Loop
```

**Step 6: ProcessMassResults()** — Parse AutoCAD log

```vb
' Reads the log file line by line
' Extracts: CutterPocket (from layer name), Area, Centroid X, Centroid Y
' NOTE: CX = -1 * CX  (centroid X is NEGATED back to positive radial)
' Sorts by Centroid_X ascending (inner radius → outer radius)
' Places results on "MassProp" sheet starting at row 9
```

### 7.4 Ellipse Geometry (Ellipse.cls)

**Ellipse parameterization** on the revolved Z plane:

```
Center: (Xc, Yc) where Xc = -radial, Yc = Z_drilling
Major axis radius: Major = PDC_radius (half the pocket diameter)
Minor axis radius: Minor = Major × cos(Rake)
Tilt angle: from DegTilt column (negated)
```

**Point on ellipse at angle θ** (parametric):

```
X_raw = (Major × Minor) / sqrt(Minor² + Major² × tan²(θ))  [adjusted for quadrant]
Y_raw = X_raw × tan(θ)

Rotated by tilt angle:
X_final = X_raw × cos(tilt) - Y_raw × sin(tilt) + Xc
Y_final = X_raw × sin(tilt) + Y_raw × cos(tilt) + Yc
```

**Point-in-ellipse test** (for masking):

```
cosT = cos(tilt_rad)
sinT = sin(tilt_rad)
eq = ((px-cx)*cosT + (py-cy)*sinT)² / Major² + ((px-cx)*sinT - (py-cy)*cosT)² / Minor²
isInside = (eq <= 1)
```

**RemovePointsInEllipse** additional conditions:
- Points are also removed if `RevolutionNum ≠ 2` (only Rev 2 retains points)
- Points beyond the gage line (x < -GageLine) are removed

**MidThetaPoint** — finds the midpoint of the visible arc (for BOUNDARY seed point):
- Tracks LowestTheta and HighestTheta of remaining visible points
- MidTheta = (LowestTheta + HighestTheta) / 2
- Returns the point on the ellipse at that angle, slightly undersized

### 7.5 Solid Edge Method (SE_Cutlets.bas / SE_Cutlets_v2.bas)

This method is for **visualization**, not quantitative measurement. It:

1. Reads coordinate system data from columns AJ-AQ (row 164+):
   - AJ: New CS name (cutter name)
   - AK: Reference CS name (parent in the coordinate system tree)
   - AL-AN: X/Y/Z offset from parent (inches → converted to meters)
   - AO-AQ: X/Y/Z rotation from parent (degrees → radians)
   - AR: Part file name (solid model)
   - AV: RGB color

2. Creates coordinate systems in Solid Edge relative to parent systems
3. For non-reference systems, imports the actual 3D cutter solid model
4. Colors each cutter body by blade

**v2 additions:**
- Unions all bodies into one design body
- Creates a gauge trim (cuts at bit radius)
- Adjusts lighting for better visualization
- Progress bar

The parent-child hierarchy creates the spatial arrangement of cutters. The "reference"
coordinate systems are intermediate positioning nodes that aren't rendered.

### 7.6 Backup Engagement Calculation (Ellipse_Min_Dist_Calc.bas)

This is a SEPARATE calculation from the cutlet plot. It determines at what ft/hr
each backup cutter begins to engage its primary.

```
For each primary-backup pair (data from columns IR-JA):
  - Primary ellipse: center, major, minor, tilt (from IR-IV)
  - Secondary ellipse: center, major, minor, tilt (from IW-JA)
  - Degrees trailing: from column IN

  Starting at 0 ft/hr, incrementing by 10 ft/hr:
    - Generate points on primary ellipse arc (90° ± 60° range, 0.5° steps)
    - Generate points on secondary ellipse arc (same range)
    - For secondary: shift Y center by helical offset = ((ft_hr*12/60)/RPM/360) × DegTrailing
    - Find minimum distance between any point pair
    - When minDist < 0.00165" (half of 0.0033), declare engagement
    - Record the ft/hr at engagement → column IQ
```

This is the module that computes the engagement speed values used for the
engagement threshold and progression metrics.

---

## 8. COMPUTATIONAL REPLICATION PLAN

### 8.1 What We Now Know

From the macro analysis, the AutoCAD cutlet plot method is a well-defined
2D computational geometry problem:

1. **Inputs**: For each cutter — radial position (X), Z position with helical offset,
   PDC diameter, tilt angle, back rake angle
2. **Ellipse construction**: Major = PDC radius, Minor = Major × cos(rake), tilted by tilt angle
3. **Three revolutions**: Original, shifted by IPR, shifted by 2×IPR
4. **Sequencing**: Cutters ordered by blade (1→6), inner to outer within each blade.
   Lower indices mask higher indices. Rev 1 masks Rev 2 which masks Rev 3.
5. **Cutlet**: For each Rev 2 cutter, the portion of its ellipse NOT covered by
   any lower-indexed ellipse (from any revolution)
6. **Output**: Centroid (X, Y) and Area for each cutlet

### 8.2 Implementation Steps

**Phase 1: Ellipse-based cutlet computation (Python)**

1. Read cutter data from Assy.Model sheet (same columns as the macros use)
2. Construct ellipses with same parameterization as Ellipse.cls
3. Create 3 revolutions (same as AdditionalEllipses)
4. For each Rev 2 ellipse, compute the non-overlapped region:
   - Use Shapely polygons (approximate ellipses as high-resolution polygons)
   - Subtract all lower-indexed ellipses from the current one
   - Clip to gauge line
   - Compute area and centroid of the resulting polygon
5. Validate against MassProp sheet values

**Phase 2: Force computation**

With centroid + area, the rest is pure math (already understood from section 4).

**Phase 3: Scoring model improvement**

Replace proxy-based scoring with actual force-distribution-based scoring.

### 8.3 Validation Strategy

- Compare computed centroids and areas against the 26 values on the MassProp sheet
- User noted these values are "close but not 100% trusted" — expect reasonable
  agreement (~5-10% tolerance) rather than exact match

---

## 9. CHAMFER EFFECTS ON CUTLET EFFICIENCY AND DURABILITY

### 9.1 The Chamfer Geometry

All cutters in the current designs have **0.020" x 45° chamfers** on the PDC (Polycrystalline
Diamond Compact) cutting face. The chamfer is a beveled ring around the cutter edge.

- The **flat face** of the cutter (inside the chamfer ring) does the efficient shearing of rock
- The **chamfer portion** does NOT cut as efficiently — it puts the PDC into compression
  rather than shearing
- The chamfer's purpose: improve cutter toughness by keeping the PDC in compression,
  preventing the diamond table from overloading and shearing off

### 9.2 Chamfer Fraction vs Cutlet Size

For a given cutter with chamfer width `w = 0.020"`:

- **Large cutlets** (high IPR, fewer cutters at a radius): The chamfer ring is a small
  fraction of the total cutlet area → most of the cutting is done by the efficient flat face
- **Small cutlets** (low IPR, many cutters at a radius): The chamfer ring is a LARGE
  fraction of the cutlet area → more of the cutting energy goes into compression
  (keeping the cutter intact) rather than shearing rock

This is a fundamental trade-off:
- **Smaller chamfer** = more aggressive, more efficient, but less tough
- **Bigger chamfer** = less aggressive, less efficient, but more tough

### 9.3 Constant Volume Principle

No matter how many cutters are doing work at a given IPR for a given hole size, **the
volume of material being removed per revolution remains constant** (it's just π × r² × IPR).

This means:
- More cutlets = smaller average cutlet areas (same volume distributed across more cutters)
- Smaller average cutlet areas = larger percentage of each cutlet is chamfer
- Larger chamfer fraction = decreased cutting efficiency (more energy in compression,
  less in shearing)
- BUT: larger chamfer fraction = increased cutter toughness (durability)

### 9.4 Implications for Scoring

**For Layout Durability (0-9):**
- Higher chamfer fraction = MORE durable (cutters kept in compression longer)
- But this is a natural consequence of having more cutters / smaller cutlets
- The force DISTRIBUTION pattern matters (how evenly forces spread across cutters),
  not the absolute force magnitudes
- Each bit size is its own bucket: a 6.5" bit and a 12.25" bit have fundamentally
  different volumes removed, so they can't be directly compared on absolute metrics
- Score within each size bucket, then scale buckets to fit the 0-9 range

**For Steerability (0-9):**
- Chamfer fraction affects gauge region particularly: gauge cutlets with high chamfer
  fraction = less aggressive lateral cutting = easier to steer
- But also: more gauge cutlets (even if small) = more wall contact = harder to steer

### 9.5 Computing Chamfer Fraction

For each cutlet with area `A` and perimeter `P`:
- Approximate chamfer area ≈ `P × w_proj` where `w_proj ≈ 0.020"` (projected chamfer width)
- Chamfer fraction ≈ `(P × 0.020) / A`
- For a circle of radius r: chamfer_frac = `2 × 0.020 / r` → inversely proportional to size
- For small cutlets, this can approach or exceed 1.0 (entire cutlet is chamfer)

Simpler aggregate metric: `chamfer_efficiency = 0.020 / sqrt(mean_cutlet_area)`
This captures the key physics: smaller cutlets → higher chamfer fraction → less efficient.

---

## 10. SCORING METHODOLOGY — PER-SIZE BUCKETS

### 10.1 Why Per-Size Buckets

Each bit size drills a different hole diameter. The volume of rock removed per revolution
is `V = π × (hole_radius)² × IPR`. A 12.25" bit removes ~3.5x the volume of a 6.5" bit.

This means:
- Absolute cutlet area totals are NOT comparable across sizes
- A 6.5" bit with total area 0.5 in² is NOT "less durable" than a 12.25" bit with 1.2 in²
- The DISTRIBUTION of areas within each size group is what matters

### 10.2 Scoring Approach

1. **Group bits by hole size** (6.5", 7.875", 8.5", 8.75", etc.)
2. **Score within each group** using distribution-based metrics:
   - Force distribution shape (normalized area distribution across radial profile)
   - Chamfer fraction (computed from cutlet areas and perimeters)
   - Load balance (Gini, blade CV — these are already size-independent)
3. **Scale groups** to fit the 0-9 range:
   - Each group gets its own internal 0-1 normalization
   - Then groups are aligned so the overall 0-9 scale is meaningful

### 10.3 Force Distribution Shape (Not Magnitudes)

Since formation yield strength is held constant at 29,500 psi, the absolute force
numbers just scale with area. What matters is the SHAPE of the distribution:
- Normalize each bit's cutlet areas to sum to 1.0
- Compare the resulting probability distribution across the radial profile
- Uniform distributions = durable (no single cutter overloaded)
- Spiky distributions = aggressive but fragile

---

## 11. ANSWERED QUESTIONS (from user)

1. **IDS** = Instantaneous Drill Speed. Formula: `IDS = ((ft/hr × 12) / 60) / RPM`
2. **Formation yield strength**: Kept constant (29,500 psi) across all assemblies
   to expose the distribution of forces, not absolute values
3. **Cutter sequencing**: CRITICALLY IMPORTANT — reviewed macros (see section 7.3)
4. **Backup engagement ellipse params (IR-JA)**: Don't use these; they're for the
   separate engagement calculation, not for cutlet geometry
5. **MassProp values**: Close but not 100% trusted
6. **Standard IPR**: 0.25 in/rev for all analyses
7. **Chamfer**: All cutters use 0.020" x 45° chamfer. Chamfer portion of cutlet
   does not cut efficiently — puts PDC in compression instead of shearing rock.
   Smaller cutlets = higher chamfer fraction = less efficient but more durable.
8. **Constant volume**: Regardless of cutter count, volume removed is constant for
   a given hole size and IPR. More cutlets = smaller avg area = higher chamfer %.
9. **Per-size scoring**: Each bit size is its own scoring bucket. Score within bucket,
   then scale to fit the 0-9 range across all sizes.
