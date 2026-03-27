"""
Building Block 2: Cuttlet Plots

Generates cuttlet cross-section plots from a Min Engagement source file.
Saves plot images to the assembly's cuttlet_plots/ folder.

Usage:
    python skills/bb2_cuttlet_plots.py <source_xlsm_path> <assy_name>

Example:
    python skills/bb2_cuttlet_plots.py "assemblies/Assy_2176r06/source/2176r06 Min Engagement v8.01.xlsm" Assy_2176r06
"""

import sys
import os
import math
import numpy as np

# Add project root to path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

from cutlet_engine import read_cutter_data_from_me, compute_cutlets, make_ellipse_polygon

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from shapely.geometry import box

# Blade colors from ME color mapping (RGB)
# Row 1 = primary (saturated), Row 2+ = secondary (lighter)
BLADE_COLORS_PRIMARY = {
    1: (255, 127, 0),   2: (204, 204, 0),   3: (0, 204, 0),
    4: (0, 204, 204),   5: (0, 102, 204),   6: (127, 0, 255),
    7: (204, 0, 204),   8: (204, 0, 0),     9: (127, 63, 0),
}
BLADE_COLORS_SECONDARY = {
    1: (255, 188, 121), 2: (255, 255, 91),  3: (97, 255, 97),
    4: (91, 255, 255),  5: (121, 188, 255), 6: (188, 121, 255),
    7: (255, 91, 255),  8: (255, 97, 97),   9: (167, 103, 40),
}

def _rgb(r, g, b):
    return (r / 255, g / 255, b / 255)

def blade_color(blade, row):
    if row == 1:
        rgb = BLADE_COLORS_PRIMARY.get(blade, (51, 51, 51))
    else:
        rgb = BLADE_COLORS_SECONDARY.get(blade, (153, 153, 153))
    return _rgb(*rgb)


def get_blade_from_name(name):
    parts = str(name).split('.')
    return int(parts[0]) if len(parts) == 2 else 1


def get_row_from_name(name):
    parts = str(name).split('.')
    if len(parts) == 2 and len(parts[1]) >= 1:
        return int(parts[1][0])
    return 1


def build_cutlet_polygons(cutters, ipr, gage_radius):
    """Build ellipses, perform subtraction, return list of (name, cutlet_polygon) pairs."""
    N = len(cutters)
    all_ellipses = []

    # Revolution 1
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

    # Revolution 2
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

    total = len(all_ellipses)
    cutlet_polys = []

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

        cutlet_polys.append((all_ellipses[idx]['name'], cutlet))

    return cutlet_polys


def extract_assy_label(assy_name):
    """Convert 'Assy_2176r06' to '2176r06' for plot title."""
    return assy_name.replace("Assy_", "")


def plot_cutlets(cutlet_polys, gage_radius, assy_name, ipr, output_path):
    """Plot all cutlet polygons colored by blade, save to output_path."""
    fig, ax = plt.subplots(1, 1, figsize=(12, 8))

    for name, cutlet in cutlet_polys:
        blade_num = get_blade_from_name(name)
        row = get_row_from_name(name)
        color = blade_color(blade_num, row)
        edgecolor = 'black' if row == 1 else '#666666'
        linewidth = 0.5 if row == 1 else 0.3

        def fill_poly(geom, c=color, ec=edgecolor, lw=linewidth):
            xs, ys = geom.exterior.xy
            # Plot at native negative-X positions (mirrored left of Y axis)
            ax.fill(list(xs), ys, color=c, edgecolor=ec, linewidth=lw)

        if cutlet.geom_type == 'Polygon':
            fill_poly(cutlet)
        elif cutlet.geom_type == 'MultiPolygon':
            for geom in cutlet.geoms:
                fill_poly(geom)

    # Gage line on the mirrored side
    ax.axvline(x=-gage_radius, color='gray', linestyle='--', linewidth=0.5, alpha=0.5)

    # Title: "2176r06 - Cutlet Plot"
    label = extract_assy_label(assy_name)
    ax.set_title(f'{label} - Cutlet Plot\nIPR = {ipr:.3f} in/rev', fontsize=12, fontweight='bold')
    ax.set_xlabel('Radial Distance (in)')
    ax.set_ylabel('Axial Distance (in)')
    ax.set_aspect('equal')
    ax.grid(True, alpha=0.2)

    # Make X tick labels show positive values even though geometry is negative
    from matplotlib.ticker import FuncFormatter
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, _: f'{abs(x):.1f}'))

    # Legend with primary colors
    blades_present = sorted(set(get_blade_from_name(name) for name, _ in cutlet_polys))
    handles = [mpatches.Patch(color=blade_color(b, 1), label=f'Blade {b}') for b in blades_present]
    ax.legend(handles=handles, loc='lower left', fontsize=9)

    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close(fig)
    print(f"  Saved: {output_path}")


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    source_path = sys.argv[1]
    assy_name = sys.argv[2]

    if not os.path.exists(source_path):
        print(f"Error: Source file not found: {source_path}")
        sys.exit(1)

    # Output directory
    output_dir = os.path.join(PROJECT_ROOT, "assemblies", assy_name, "cuttlet_plots")
    os.makedirs(output_dir, exist_ok=True)

    print(f"BB2: Cuttlet Plots — {assy_name}")
    print(f"  Source: {source_path}")

    # Read cutter data
    cutters, ipr, gage_radius = read_cutter_data_from_me(source_path)
    print(f"  Cutters: {len(cutters)}, IPR: {ipr:.3f}, Gage: {gage_radius:.4f}")

    # Build cutlet polygons
    print("  Computing cutlet polygons...")
    cutlet_polys = build_cutlet_polygons(cutters, ipr, gage_radius)
    print(f"  Cutlets: {len(cutlet_polys)}")

    # Generate plot
    output_file = os.path.join(output_dir, f"{assy_name}_cuttlet_plot.png")
    plot_cutlets(cutlet_polys, gage_radius, assy_name, ipr, output_file)

    # Summary stats
    total_area = sum(c.area for _, c in cutlet_polys)
    print(f"\n  Total cutlet area: {total_area:.4f} in²")
    print(f"  Blade breakdown:")
    blade_areas = {}
    for name, cutlet in cutlet_polys:
        blade = get_blade_from_name(name)
        blade_areas[blade] = blade_areas.get(blade, 0) + cutlet.area
    for b in sorted(blade_areas):
        print(f"    Blade {b}: {blade_areas[b]:.4f} in²")

    print(f"\nBB2 complete for {assy_name}.")


if __name__ == '__main__':
    main()
