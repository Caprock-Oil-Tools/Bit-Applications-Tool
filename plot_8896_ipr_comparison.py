"""
Generate cutlet plots for 8896r00 at three IPR values (0.100, 0.250, 0.500 in/rev)
to study F-type behavior across drilling speeds.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from cutlet_engine import read_cutter_data_from_me, compute_cutlets, make_ellipse_polygon
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.collections import PatchCollection
from shapely.geometry import box
import math
import numpy as np

# Blade colors (matching existing plot style)
BLADE_COLORS = {
    1: '#e41a1c', 2: '#377eb8', 3: '#4daf4a',
    4: '#ff7f00', 5: '#984ea3', 6: '#a65628',
    7: '#f781bf', 8: '#999999', 9: '#66c2a5',
}

def get_blade_from_name(name):
    parts = str(name).split('.')
    return int(parts[0]) if len(parts) == 2 else 1

def get_row_from_name(name):
    parts = str(name).split('.')
    if len(parts) == 2 and len(parts[1]) >= 1:
        return int(parts[1][0])
    return 1

def plot_cutlets_for_file(ax, filepath, label, ipr_override=None):
    """Generate a cutlet plot on the given axes."""
    cutters, ipr_file, gage_radius = read_cutter_data_from_me(filepath)
    ipr = ipr_override if ipr_override is not None else ipr_file
    print(f"\n{label}: {len(cutters)} cutters, IPR={ipr:.3f} (file={ipr_file:.3f}), gage_radius={gage_radius:.4f}")

    # Compute cutlets
    cutlet_results = compute_cutlets(cutters, ipr, gage_radius)
    cutlet_results = [r for r in cutlet_results if r['area'] >= 0.001]

    # Also need the actual polygon shapes for plotting
    # Rebuild polygons to draw them
    N = len(cutters)
    all_ellipses = []
    for c in cutters:
        xc = -c['radial']
        yc = c['z_drill']
        major = c['major']
        tilt = -c['tilt']
        rake = c['rake']
        minor = major * math.cos(math.radians(rake))
        all_ellipses.append({
            'name': c['name'], 'xc': xc, 'yc': yc,
            'major': major, 'minor': minor, 'tilt': tilt,
            'rev': 1, 'massprop': False,
        })

    for i in range(N - 1):
        e = all_ellipses[i]
        all_ellipses.append({
            'name': e['name'], 'xc': e['xc'],
            'yc': e['yc'] + ipr,
            'major': e['major'], 'minor': e['minor'],
            'tilt': e['tilt'], 'rev': 2, 'massprop': True,
        })

    polygons = []
    for e in all_ellipses:
        poly = make_ellipse_polygon(e['xc'], e['yc'], e['major'], e['minor'], e['tilt'], n_points=720)
        polygons.append(poly)

    all_y = [e['yc'] for e in all_ellipses]
    y_min_clip = min(all_y) - 2.0
    y_max_clip = max(all_y) + 2.0
    clip_box = box(-gage_radius, y_min_clip, 0, y_max_clip)

    # Compute and draw each cutlet polygon with blade color
    total = len(all_ellipses)
    row1_areas_by_blade = {}
    row2_areas_by_blade = {}

    for idx in range(total):
        if not all_ellipses[idx]['massprop']:
            continue

        cutlet = polygons[idx]
        for mask_idx in range(idx):
            if not cutlet.is_empty:
                cutlet = cutlet.difference(polygons[mask_idx])
        if not cutlet.is_empty:
            cutlet = cutlet.intersection(clip_box)
        if cutlet.is_empty or cutlet.area < 1e-10:
            continue

        name = all_ellipses[idx]['name']
        blade = get_blade_from_name(name)
        row = get_row_from_name(name)
        color = BLADE_COLORS.get(blade, '#333333')

        # Track areas by blade and row
        if row == 1:
            row1_areas_by_blade.setdefault(blade, []).append(cutlet.area)
        else:
            row2_areas_by_blade.setdefault(blade, []).append(cutlet.area)

        # Draw the cutlet polygon
        alpha = 0.7 if row == 1 else 0.35
        edgecolor = 'black' if row == 1 else '#666666'
        linewidth = 0.5 if row == 1 else 0.3

        if cutlet.geom_type == 'Polygon':
            xs, ys = cutlet.exterior.xy
            ax.fill([-x for x in xs], ys, color=color, alpha=alpha,
                    edgecolor=edgecolor, linewidth=linewidth)
        elif cutlet.geom_type == 'MultiPolygon':
            for geom in cutlet.geoms:
                xs, ys = geom.exterior.xy
                ax.fill([-x for x in xs], ys, color=color, alpha=alpha,
                        edgecolor=edgecolor, linewidth=linewidth)

    # Draw gauge line
    ax.axvline(x=gage_radius, color='gray', linestyle='--', linewidth=0.5, alpha=0.5)

    ax.set_title(f'{label}\nIPR = {ipr:.3f} in/rev', fontsize=10, fontweight='bold')
    ax.set_xlabel('Radial Position (in)')
    ax.set_ylabel('Z (in)')
    ax.set_aspect('equal')
    ax.grid(True, alpha=0.2)

    # Print per-blade area summary
    print(f"\n  Row 1 (primary) cutlet areas by blade:")
    for blade in sorted(row1_areas_by_blade.keys()):
        areas = row1_areas_by_blade[blade]
        print(f"    Blade {blade}: {len(areas)} cutlets, total={sum(areas):.4f}, "
              f"mean={np.mean(areas):.4f}, min={min(areas):.4f}, max={max(areas):.4f}")

    if row2_areas_by_blade:
        print(f"  Row 2 (backup) cutlet areas by blade:")
        for blade in sorted(row2_areas_by_blade.keys()):
            areas = row2_areas_by_blade[blade]
            print(f"    Blade {blade}: {len(areas)} cutlets, total={sum(areas):.4f}, "
                  f"mean={np.mean(areas):.4f}")

    # Print radial balance analysis
    print(f"\n  Radial primary cutlet analysis (per radial bin):")
    # Bin primary cutlets by radial position
    n_bins = 15
    bin_width = gage_radius / n_bins
    bins = {}  # bin_idx -> {blade: area}
    for r in cutlet_results:
        row = get_row_from_name(r['name'])
        if row != 1:
            continue
        blade = get_blade_from_name(r['name'])
        bin_idx = min(int(r['centroid_x'] / bin_width), n_bins - 1)
        if bin_idx not in bins:
            bins[bin_idx] = {}
        bins[bin_idx][blade] = bins[bin_idx].get(blade, 0) + r['area']

    for bi in sorted(bins.keys()):
        r_lo = bi * bin_width
        r_hi = (bi + 1) * bin_width
        blade_areas = bins[bi]
        n_blades = len(blade_areas)
        vals = list(blade_areas.values())
        if n_blades >= 2:
            cv = np.std(vals) / np.mean(vals) if np.mean(vals) > 0 else 0
            ratio = max(vals) / min(vals) if min(vals) > 0 else float('inf')
            blades_str = ', '.join(f'B{b}={a:.4f}' for b, a in sorted(blade_areas.items()))
            print(f"    [{r_lo:.2f}-{r_hi:.2f}] {n_blades} blades, CV={cv:.3f}, "
                  f"max/min={ratio:.2f} | {blades_str}")
        else:
            blades_str = ', '.join(f'B{b}={a:.4f}' for b, a in sorted(blade_areas.items()))
            print(f"    [{r_lo:.2f}-{r_hi:.2f}] {n_blades} blade  | {blades_str}")

    return cutlet_results


# --- Main ---
base = "Bit Designs/8896"
files = [
    (f"{base}/8896r00 @ 0.100 Min Engagement v8.01.xlsm", "8896r00 @ 0.100"),
    (f"{base}/8896r00 @ 0.250 Min Engagement v8.01.xlsm", "8896r00 @ 0.250"),
    (f"{base}/8896r00 @ 0.500 Min Engagement v8.01.xlsm", "8896r00 @ 0.500"),
]

fig, axes = plt.subplots(1, 3, figsize=(24, 8))

for ax, (fpath, label) in zip(axes, files):
    plot_cutlets_for_file(ax, fpath, label)

# Legend
handles = [mpatches.Patch(color=BLADE_COLORS[b], label=f'Blade {b}') for b in range(1, 7)]
handles.append(mpatches.Patch(color='gray', alpha=0.35, label='Row 2 (backup)'))
fig.legend(handles=handles, loc='lower center', ncol=7, fontsize=9)

plt.suptitle('8896r00 F-Type — Cutlet Comparison at Three IPR Values', fontsize=14, fontweight='bold')
plt.tight_layout(rect=[0, 0.06, 1, 0.95])
plt.savefig('cutlet_plots_8896_ipr_comparison.png', dpi=150, bbox_inches='tight')
print("\n\nSaved: cutlet_plots_8896_ipr_comparison.png")
