"""
Compute Layout Durability (0-9) and Steerability (0-9) ratings for each bit design
by analyzing cutter pocket positions and orientations from Min Engagement files.

Durability 0 = most aggressive, 9 = most durable
Steerability 0 = least steerable, 9 = most steerable
"""

import openpyxl
import os
import glob
import math
import json
import warnings
import numpy as np
from collections import defaultdict

warnings.filterwarnings("ignore", category=UserWarning)

WORKBOOK_PATH = "Bit Applications Tool r2.xlsx"
BIT_DESIGNS_DIR = "Bit Designs"


def extract_cutters_from_me_file(filepath):
    """Extract cutter position and orientation data from a Min Engagement file."""
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb["Assy.Model"]

    # Detect version from filename
    basename = os.path.basename(filepath)
    is_v6 = "v6." in basename

    # Get bit diameter from row 1 or Settings sheet
    bit_diameter = None
    for row_idx, row in enumerate(ws.iter_rows(values_only=False), 1):
        if row_idx == 1:
            for c in row:
                if hasattr(c, "column_letter") and c.column_letter == "C" and c.value:
                    val = str(c.value)
                    # Extract leading number as bit diameter
                    parts = val.split()
                    if parts:
                        try:
                            bit_diameter = float(parts[0])
                        except ValueError:
                            pass
            break

    # Also try Settings sheet for bit diameter
    if bit_diameter is None:
        try:
            ws_settings = wb["Settings"]
            for row_idx, row in enumerate(ws_settings.iter_rows(values_only=False), 1):
                if row_idx == 19:  # Row where "Bit ø" typically is
                    for c in row:
                        if (
                            hasattr(c, "column_letter")
                            and c.column_letter == "C"
                            and isinstance(c.value, (int, float))
                        ):
                            bit_diameter = float(c.value)
                            break
                if bit_diameter:
                    break
        except (KeyError, Exception):
            pass

    cutters = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=False), 1):
        if row_idx <= 8:
            continue

        data = {}
        for c in row:
            if not hasattr(c, "column_letter"):
                continue
            cl = c.column_letter
            if c.value is None:
                continue
            if cl == "B":
                data["x"] = c.value
            elif cl == "C":
                data["y"] = c.value
            elif cl == "D":
                data["z"] = c.value
            elif cl == "E":
                data["dx"] = c.value
            elif cl == "F":
                data["dy"] = c.value
            elif cl == "G":
                data["dz"] = c.value
            elif cl == "H":
                data["pocket_radius"] = c.value
            elif cl == "I":
                data["pocket_depth"] = c.value
            elif cl == "J":
                data["name"] = c.value
            elif cl == "K":
                data["element"] = c.value
            elif cl == "L":
                data["zone"] = c.value
            elif cl == "R":
                data["radial_pos"] = c.value
            elif cl == "Q":
                data["theta_deg"] = c.value
            elif cl == "AN":
                data["tilt_raw"] = c.value
            elif cl == "AO":
                data["backrake_raw"] = c.value
            elif cl == "AP":
                data["siderake_raw"] = c.value

        # Filter: must have a name (cutter ID) and position
        if data.get("name") is None or data.get("x") is None:
            continue

        # Skip "Flip Vector?" or header-like entries
        name_str = str(data["name"])
        if "flip" in name_str.lower() or "vector" in name_str.lower():
            continue
        if "Part" in name_str:
            continue

        # Parse blade number from name (e.g., "1.113" -> blade 1)
        if "." in name_str:
            blade_str = name_str.split(".")[0]
            try:
                data["blade"] = int(blade_str)
            except ValueError:
                continue
        else:
            continue

        # Convert back rake for v6.xx files
        if is_v6 and data.get("backrake_raw") is not None:
            raw = data["backrake_raw"]
            if isinstance(raw, (int, float)) and raw > 90:
                data["backrake"] = 180.0 - raw
            elif isinstance(raw, (int, float)) and raw == 0:
                data["backrake"] = None  # center cutter, skip
            else:
                data["backrake"] = raw if isinstance(raw, (int, float)) else None
        elif data.get("backrake_raw") is not None:
            raw = data["backrake_raw"]
            data["backrake"] = raw if isinstance(raw, (int, float)) else None
        else:
            data["backrake"] = None

        # Side rake
        if data.get("siderake_raw") is not None:
            raw = data["siderake_raw"]
            if is_v6 and isinstance(raw, (int, float)):
                # v6 side rake is also in raw angle form
                data["siderake"] = abs(180.0 - raw) if raw > 90 else raw
            else:
                data["siderake"] = raw if isinstance(raw, (int, float)) else None
        else:
            data["siderake"] = None

        # Compute radial position if not available
        if data.get("radial_pos") is None and data.get("x") is not None and data.get("y") is not None:
            x, y = data["x"], data["y"]
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                data["radial_pos"] = math.sqrt(x**2 + y**2)

        cutters.append(data)

    wb.close()
    return cutters, bit_diameter


def compute_metrics(cutters, bit_diameter, layout_type, pattern_order, blade_count_meta):
    """Compute durability and steerability metrics from cutter data."""
    if not cutters:
        return None, None

    # Organize by blade
    blades = defaultdict(list)
    for c in cutters:
        blades[c["blade"]].append(c)

    num_blades = len(blades)
    if num_blades == 0:
        return None, None

    bit_radius = bit_diameter / 2.0 if bit_diameter else None

    # --- DURABILITY METRICS ---

    # 1. Total cutter count (more cutters = more durable)
    total_cutters = len(cutters)

    # 2. Average cutters per blade
    avg_cutters_per_blade = total_cutters / num_blades

    # 3. Back rake statistics (higher back rake = more durable/conservative)
    all_backrakes = [c["backrake"] for c in cutters if c.get("backrake") is not None
                     and isinstance(c["backrake"], (int, float)) and 0 < c["backrake"] < 60]
    avg_backrake = np.mean(all_backrakes) if all_backrakes else 20.0
    min_backrake = min(all_backrakes) if all_backrakes else 15.0

    # 4. Radial coverage and redundancy
    # For each radial zone, count how many blades have cutters there
    if bit_radius:
        gauge_radius = bit_radius
    else:
        # Estimate from max radial position
        all_radii = [abs(c.get("radial_pos", 0)) for c in cutters
                     if isinstance(c.get("radial_pos"), (int, float))]
        gauge_radius = max(all_radii) if all_radii else 4.0

    # Divide the profile into zones based on percentage of gauge radius
    num_zones = 10
    zone_width = gauge_radius / num_zones
    zone_blade_coverage = defaultdict(set)  # zone -> set of blade numbers

    for c in cutters:
        rp = c.get("radial_pos")
        if rp is None or not isinstance(rp, (int, float)):
            continue
        rp = abs(rp)
        zone_idx = min(int(rp / zone_width), num_zones - 1) if zone_width > 0 else 0
        zone_blade_coverage[zone_idx].add(c["blade"])

    # Average blades per zone (higher = more redundant = more durable)
    if zone_blade_coverage:
        avg_blades_per_zone = np.mean([len(v) for v in zone_blade_coverage.values()])
    else:
        avg_blades_per_zone = 1.0

    # 5. Cutter spacing uniformity in radial direction per blade
    spacing_scores = []
    for blade_num, blade_cutters in blades.items():
        radii = sorted([abs(c.get("radial_pos", 0)) for c in blade_cutters
                        if isinstance(c.get("radial_pos"), (int, float))])
        if len(radii) > 1:
            spacings = [radii[i+1] - radii[i] for i in range(len(radii)-1)]
            if spacings:
                mean_sp = np.mean(spacings)
                std_sp = np.std(spacings)
                # Coefficient of variation: lower = more uniform = more durable
                cv = std_sp / mean_sp if mean_sp > 0 else 1.0
                spacing_scores.append(cv)

    avg_spacing_cv = np.mean(spacing_scores) if spacing_scores else 0.5

    # 6. Layout type scoring
    layout_score = 0.0
    if layout_type:
        lt = layout_type.lower()
        if "redundant" in lt:
            layout_score = 1.0  # Most durable layout
        elif "singleset" in lt or "single" in lt:
            layout_score = 0.4
        else:
            layout_score = 0.5

    # 7. Pocket depth uniformity (deeper pockets = more secure = more durable)
    depths = [c.get("pocket_depth") for c in cutters
              if isinstance(c.get("pocket_depth"), (int, float))]
    avg_depth = np.mean(depths) if depths else 0.4

    # 8. Cutter density (cutters per unit of radial coverage)
    if gauge_radius > 0:
        cutter_density = total_cutters / gauge_radius
    else:
        cutter_density = total_cutters / 4.0

    # --- STEERABILITY METRICS ---

    # 1. Gauge region analysis
    # Cutters near gauge (>85% of gauge radius) affect steerability
    gauge_threshold = 0.85 * gauge_radius
    gauge_cutters = [c for c in cutters
                     if isinstance(c.get("radial_pos"), (int, float))
                     and abs(c.get("radial_pos", 0)) > gauge_threshold]
    gauge_cutter_ratio = len(gauge_cutters) / total_cutters if total_cutters > 0 else 0

    # 2. Gauge cutter back rake (higher at gauge = less steerable)
    gauge_backrakes = [c["backrake"] for c in gauge_cutters
                       if c.get("backrake") is not None
                       and isinstance(c["backrake"], (int, float)) and 0 < c["backrake"] < 60]
    avg_gauge_backrake = np.mean(gauge_backrakes) if gauge_backrakes else 30.0

    # 3. Profile aggressiveness from cutter Z positions
    # More variation in Z = more aggressive profile = potentially more steerable
    all_z = [c.get("z") for c in cutters if isinstance(c.get("z"), (int, float))]
    z_range = (max(all_z) - min(all_z)) if len(all_z) > 1 else 0

    # 4. Side rake analysis (more side rake = more steerable)
    all_siderakes = [c.get("siderake") for c in cutters
                     if c.get("siderake") is not None
                     and isinstance(c["siderake"], (int, float)) and abs(c["siderake"]) < 30]
    avg_siderake = np.mean([abs(s) for s in all_siderakes]) if all_siderakes else 0

    # 5. Force imbalance estimation
    # Compute approximate lateral force vectors from cutter orientations
    fx_sum = 0.0
    fy_sum = 0.0
    force_count = 0
    for c in cutters:
        dx = c.get("dx")
        dy = c.get("dy")
        if isinstance(dx, (int, float)) and isinstance(dy, (int, float)):
            fx_sum += dx
            fy_sum += dy
            force_count += 1

    if force_count > 0:
        # Normalized resultant lateral force (higher = more steerable tendency)
        lateral_resultant = math.sqrt(fx_sum**2 + fy_sum**2) / force_count
    else:
        lateral_resultant = 0

    # 6. Blade count effect (fewer blades can be more steerable)
    blade_count_for_steer = num_blades

    # 7. Cone region aggressiveness
    # Cutters in inner 30% of radius with low backrake = aggressive cone = more steerable
    cone_threshold = 0.30 * gauge_radius
    cone_cutters = [c for c in cutters
                    if isinstance(c.get("radial_pos"), (int, float))
                    and abs(c.get("radial_pos", 0)) < cone_threshold]
    cone_backrakes = [c["backrake"] for c in cone_cutters
                      if c.get("backrake") is not None
                      and isinstance(c["backrake"], (int, float)) and 0 < c["backrake"] < 60]
    avg_cone_backrake = np.mean(cone_backrakes) if cone_backrakes else 15.0

    # 8. Angular spread of cutters (more uniform angular distribution = less steerable)
    all_thetas = [c.get("theta_deg") for c in cutters
                  if isinstance(c.get("theta_deg"), (int, float))]
    if len(all_thetas) > 2:
        sorted_thetas = sorted([t % 360 for t in all_thetas])
        angular_gaps = []
        for i in range(len(sorted_thetas) - 1):
            angular_gaps.append(sorted_thetas[i+1] - sorted_thetas[i])
        angular_gaps.append(360 - sorted_thetas[-1] + sorted_thetas[0])
        max_angular_gap = max(angular_gaps) if angular_gaps else 60
    else:
        max_angular_gap = 60

    return {
        "total_cutters": total_cutters,
        "num_blades": num_blades,
        "avg_cutters_per_blade": avg_cutters_per_blade,
        "avg_backrake": avg_backrake,
        "min_backrake": min_backrake,
        "avg_blades_per_zone": avg_blades_per_zone,
        "avg_spacing_cv": avg_spacing_cv,
        "layout_score": layout_score,
        "avg_depth": avg_depth,
        "cutter_density": cutter_density,
        "gauge_cutter_ratio": gauge_cutter_ratio,
        "avg_gauge_backrake": avg_gauge_backrake,
        "z_range": z_range,
        "avg_siderake": avg_siderake,
        "lateral_resultant": lateral_resultant,
        "blade_count_for_steer": blade_count_for_steer,
        "avg_cone_backrake": avg_cone_backrake,
        "max_angular_gap": max_angular_gap,
        "gauge_radius": gauge_radius,
    }, None


def normalize_to_scale(values, low=0, high=9, invert=False):
    """Normalize a list of values to 0-9 scale. invert=True means higher input -> lower output."""
    arr = np.array(values, dtype=float)
    valid = ~np.isnan(arr)
    if valid.sum() < 2:
        return [4.5] * len(values)

    vmin = np.nanmin(arr)
    vmax = np.nanmax(arr)

    if vmax == vmin:
        return [4.5] * len(values)

    normalized = (arr - vmin) / (vmax - vmin)  # 0 to 1
    if invert:
        normalized = 1.0 - normalized

    scaled = normalized * (high - low) + low
    return [round(float(s), 1) if not np.isnan(s) else None for s in scaled]


def main():
    # Load main workbook to get bit metadata
    print("Loading main workbook...")
    wb_main = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True)
    ws = wb_main["Sheet1"]

    bits = []
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=False):
        bit_data = {}
        for c in row:
            cl = c.column_letter
            if cl == "B":
                bit_data["bit_num"] = c.value
            elif cl == "X":
                bit_data["size"] = c.value
            elif cl == "AA":
                bit_data["blade_count"] = c.value
            elif cl == "AR":
                bit_data["layout_type"] = c.value
            elif cl == "AS":
                bit_data["pattern_order"] = c.value
            elif cl == "AO":
                bit_data["durability"] = c.value
            elif cl == "AP":
                bit_data["steerability"] = c.value
        if bit_data.get("bit_num") is not None:
            bit_data["row_num"] = row[0].row
            bits.append(bit_data)

    print(f"Found {len(bits)} bits in workbook")

    # Find Min Engagement files for each bit
    all_me_files = glob.glob(f"{BIT_DESIGNS_DIR}/**/*Min Engagement*.xlsm", recursive=True)
    print(f"Found {len(all_me_files)} Min Engagement files")

    # Map bit numbers to ME files (prefer latest version)
    bit_me_map = {}
    for f in all_me_files:
        folder_parts = f.replace("\\", "/").split("/")
        folder_name = folder_parts[1] if len(folder_parts) > 1 else ""

        for bit in bits:
            bn = str(int(bit["bit_num"])) if isinstance(bit["bit_num"], (int, float)) else str(bit["bit_num"])
            # Match by folder name or filename containing bit number
            if folder_name == bn or bn in os.path.basename(f).split(" ")[0].replace(".", "").replace("r", ""):
                # Prefer highest version
                version = 0
                bname = os.path.basename(f)
                if "v8." in bname:
                    version = 8
                elif "v7." in bname:
                    version = 7
                elif "v6." in bname:
                    version = 6

                # Skip XH/variant files unless it's the only one
                is_variant = any(x in bname.upper() for x in ["XH", "RIPPER", "W "])

                current = bit_me_map.get(bn)
                if current is None:
                    bit_me_map[bn] = (f, version, is_variant)
                else:
                    _, cur_ver, cur_variant = current
                    # Prefer non-variant, then higher version
                    if is_variant and not cur_variant:
                        continue
                    if not is_variant and cur_variant:
                        bit_me_map[bn] = (f, version, is_variant)
                    elif version > cur_ver:
                        bit_me_map[bn] = (f, version, is_variant)

    matched = {bn: info[0] for bn, info in bit_me_map.items()}
    print(f"Matched {len(matched)} bits to Min Engagement files")

    # Extract metrics for all matched bits
    all_metrics = {}
    for i, bit in enumerate(bits):
        bn = str(int(bit["bit_num"])) if isinstance(bit["bit_num"], (int, float)) else str(bit["bit_num"])

        if bn not in matched:
            continue

        filepath = matched[bn]
        print(f"  [{i+1}/{len(bits)}] Processing bit {bn}: {os.path.basename(filepath)}")

        try:
            cutters, bit_diameter = extract_cutters_from_me_file(filepath)

            # Use size from main workbook if not found in ME file
            if bit_diameter is None and bit.get("size"):
                bit_diameter = float(bit["size"])

            metrics, _ = compute_metrics(
                cutters, bit_diameter,
                bit.get("layout_type"),
                bit.get("pattern_order"),
                bit.get("blade_count"),
            )

            if metrics:
                all_metrics[bn] = metrics
        except Exception as e:
            print(f"    ERROR: {e}")

    print(f"\nSuccessfully computed metrics for {len(all_metrics)} bits")

    if not all_metrics:
        print("No metrics computed. Exiting.")
        return

    # --- COMPUTE DURABILITY SCORES ---
    # Composite durability from multiple factors
    bit_numbers_ordered = [bn for bn in [
        str(int(b["bit_num"])) if isinstance(b["bit_num"], (int, float)) else str(b["bit_num"])
        for b in bits
    ] if bn in all_metrics]

    durability_components = {}
    for bn in bit_numbers_ordered:
        m = all_metrics[bn]
        # Higher = more durable for all these:
        d_backrake = m["avg_backrake"]           # Higher avg backrake = more durable
        d_density = m["cutter_density"]           # Higher density = more durable
        d_redundancy = m["avg_blades_per_zone"]   # More blades per zone = more durable
        d_layout = m["layout_score"]              # Redundant > SingleSet
        d_spacing_uniformity = 1.0 - min(m["avg_spacing_cv"], 1.0)  # Lower CV = more uniform = more durable
        d_cutters = m["avg_cutters_per_blade"]    # More cutters per blade = more durable

        durability_components[bn] = {
            "backrake": d_backrake,
            "density": d_density,
            "redundancy": d_redundancy,
            "layout": d_layout,
            "uniformity": d_spacing_uniformity,
            "cutters_per_blade": d_cutters,
        }

    # Normalize each component to 0-1, then weighted sum
    component_names = ["backrake", "density", "redundancy", "layout", "uniformity", "cutters_per_blade"]
    weights_dur = {
        "backrake": 0.20,       # Back rake is a strong durability indicator
        "density": 0.15,        # Cutter density matters
        "redundancy": 0.25,     # Radial redundancy is key for durability
        "layout": 0.20,         # Layout type (redundant vs single-set)
        "uniformity": 0.10,     # Spacing uniformity
        "cutters_per_blade": 0.10,  # More cutters = more durable
    }

    # Normalize each component
    for comp in component_names:
        vals = [durability_components[bn][comp] for bn in bit_numbers_ordered]
        vmin, vmax = min(vals), max(vals)
        for bn in bit_numbers_ordered:
            if vmax > vmin:
                durability_components[bn][comp] = (durability_components[bn][comp] - vmin) / (vmax - vmin)
            else:
                durability_components[bn][comp] = 0.5

    # Compute weighted durability score
    durability_raw = {}
    for bn in bit_numbers_ordered:
        score = sum(durability_components[bn][comp] * weights_dur[comp] for comp in component_names)
        durability_raw[bn] = score

    # Scale to 0-9
    dur_vals = [durability_raw[bn] for bn in bit_numbers_ordered]
    dur_min, dur_max = min(dur_vals), max(dur_vals)
    durability_scores = {}
    for bn in bit_numbers_ordered:
        if dur_max > dur_min:
            scaled = ((durability_raw[bn] - dur_min) / (dur_max - dur_min)) * 9.0
        else:
            scaled = 4.5
        durability_scores[bn] = round(scaled, 1)

    # --- COMPUTE STEERABILITY SCORES ---
    steerability_components = {}
    for bn in bit_numbers_ordered:
        m = all_metrics[bn]
        # For steerability, lower gauge contact and more aggressive profile = more steerable
        s_gauge_ratio = 1.0 - m["gauge_cutter_ratio"]    # Fewer gauge cutters = more steerable
        s_gauge_br = 1.0 - min(m["avg_gauge_backrake"] / 45.0, 1.0)  # Lower gauge BR = more steerable
        s_z_range = m["z_range"]                           # More Z variation = more profile = more steerable
        s_siderake = m["avg_siderake"]                     # More side rake = more steerable
        s_lateral = m["lateral_resultant"]                 # More lateral imbalance = more steerable
        s_cone_aggr = 1.0 - min(m["avg_cone_backrake"] / 30.0, 1.0)  # Lower cone BR = more aggressive = more steerable

        # Fewer blades can be more steerable (inverted)
        s_blades = 1.0 / m["blade_count_for_steer"] if m["blade_count_for_steer"] > 0 else 0.2

        steerability_components[bn] = {
            "gauge_openness": s_gauge_ratio,
            "gauge_aggressiveness": s_gauge_br,
            "profile_depth": s_z_range,
            "siderake": s_siderake,
            "lateral_force": s_lateral,
            "cone_aggressiveness": s_cone_aggr,
            "blade_factor": s_blades,
        }

    steer_comp_names = ["gauge_openness", "gauge_aggressiveness", "profile_depth",
                        "siderake", "lateral_force", "cone_aggressiveness", "blade_factor"]
    weights_steer = {
        "gauge_openness": 0.20,        # Fewer gauge cutters = more steerable
        "gauge_aggressiveness": 0.20,   # Lower gauge back rake = more steerable
        "profile_depth": 0.15,          # Deeper profile = more steerable
        "siderake": 0.10,               # Side rake affects steerability
        "lateral_force": 0.10,          # Force imbalance
        "cone_aggressiveness": 0.15,    # Aggressive cone = more steerable
        "blade_factor": 0.10,           # Fewer blades = more steerable
    }

    # Normalize each component
    for comp in steer_comp_names:
        vals = [steerability_components[bn][comp] for bn in bit_numbers_ordered]
        vmin, vmax = min(vals), max(vals)
        for bn in bit_numbers_ordered:
            if vmax > vmin:
                steerability_components[bn][comp] = (steerability_components[bn][comp] - vmin) / (vmax - vmin)
            else:
                steerability_components[bn][comp] = 0.5

    # Compute weighted steerability score
    steerability_raw = {}
    for bn in bit_numbers_ordered:
        score = sum(steerability_components[bn][comp] * weights_steer[comp] for comp in steer_comp_names)
        steerability_raw[bn] = score

    # Scale to 0-9
    steer_vals = [steerability_raw[bn] for bn in bit_numbers_ordered]
    steer_min, steer_max = min(steer_vals), max(steer_vals)
    steerability_scores = {}
    for bn in bit_numbers_ordered:
        if steer_max > steer_min:
            scaled = ((steerability_raw[bn] - steer_min) / (steer_max - steer_min)) * 9.0
        else:
            scaled = 4.5
        steerability_scores[bn] = round(scaled, 1)

    # --- PRINT RESULTS ---
    print("\n" + "=" * 80)
    print(f"{'Bit #':<8} {'Layout':<18} {'Blades':<7} {'Cutters':<9} {'Durability':<12} {'Steerability':<12}")
    print("=" * 80)
    for bit in bits:
        bn = str(int(bit["bit_num"])) if isinstance(bit["bit_num"], (int, float)) else str(bit["bit_num"])
        dur = durability_scores.get(bn, "-")
        steer = steerability_scores.get(bn, "-")
        layout = str(bit.get("layout_type", ""))[:16]
        m = all_metrics.get(bn)
        blades = m["num_blades"] if m else "-"
        cutters = m["total_cutters"] if m else "-"
        print(f"  {bn:<8} {layout:<18} {str(blades):<7} {str(cutters):<9} {str(dur):<12} {str(steer):<12}")

    # --- WRITE TO WORKBOOK ---
    print(f"\nWriting scores to {WORKBOOK_PATH}...")
    # Reload without data_only to preserve formulas
    wb_write = openpyxl.load_workbook(WORKBOOK_PATH)
    ws_write = wb_write["Sheet1"]

    updates = 0
    for row in ws_write.iter_rows(min_row=5, max_row=ws_write.max_row, values_only=False):
        bit_val = None
        ao_cell = None
        ap_cell = None
        for c in row:
            if c.column_letter == "B":
                bit_val = c.value
            elif c.column_letter == "AO":
                ao_cell = c
            elif c.column_letter == "AP":
                ap_cell = c

        if bit_val is None:
            continue

        bn = str(int(bit_val)) if isinstance(bit_val, (int, float)) else str(bit_val)

        if bn in durability_scores and ao_cell is not None:
            ao_cell.value = durability_scores[bn]
            updates += 1
        if bn in steerability_scores and ap_cell is not None:
            ap_cell.value = steerability_scores[bn]

    wb_write.save(WORKBOOK_PATH)
    print(f"Updated {updates} rows in columns AO & AP")

    # Save detailed metrics to JSON for reference
    metrics_output = {}
    for bn in bit_numbers_ordered:
        metrics_output[bn] = {
            "durability_score": durability_scores.get(bn),
            "steerability_score": steerability_scores.get(bn),
            "raw_metrics": {k: round(v, 4) if isinstance(v, float) else v
                           for k, v in all_metrics[bn].items()},
        }

    with open("bit_ratings_analysis.json", "w") as f:
        json.dump(metrics_output, f, indent=2)
    print("Saved detailed metrics to bit_ratings_analysis.json")


if __name__ == "__main__":
    main()
