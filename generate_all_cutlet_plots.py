"""
Batch generate cutlet plots for all assemblies that have source files.
Skips assemblies that already have a cutlet plot.
"""

import os
import sys
import traceback

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

from cutlet_engine import read_cutter_data_from_me
from skills.bb2_cuttlet_plots import build_cutlet_polygons, plot_cutlets

MASTER_DIR = os.path.join(PROJECT_ROOT, "Master")


def find_source_file(assy_dir):
    """Find the Min Engagement source file in an assembly folder."""
    for f in os.listdir(assy_dir):
        if "Min Engagement" in f and f.endswith(".xlsm") and not f.startswith("~$"):
            return os.path.join(assy_dir, f)
    return None


def main():
    assemblies = sorted([
        d for d in os.listdir(MASTER_DIR)
        if os.path.isdir(os.path.join(MASTER_DIR, d))
    ])

    total = len(assemblies)
    generated = 0
    skipped = 0
    failed = []

    for i, assy_name in enumerate(assemblies, 1):
        assy_dir = os.path.join(MASTER_DIR, assy_name)
        plot_file = os.path.join(assy_dir, f"{assy_name}_cutlet_plot.png")

        # Skip if already generated
        if os.path.exists(plot_file):
            skipped += 1
            continue

        source = find_source_file(assy_dir)
        if not source:
            continue

        print(f"\n[{i}/{total}] {assy_name}")
        try:
            cutters, ipr, gage_radius = read_cutter_data_from_me(source)
            cutlet_polys = build_cutlet_polygons(cutters, ipr, gage_radius)
            plot_cutlets(cutlet_polys, gage_radius, assy_name, ipr, plot_file)
            generated += 1
        except Exception as e:
            print(f"  FAILED: {e}")
            traceback.print_exc()
            failed.append((assy_name, str(e)))

    print(f"\n{'='*60}")
    print(f"Generated: {generated}")
    print(f"Skipped (already exist): {skipped}")
    print(f"Failed: {len(failed)}")
    if failed:
        print("\nFailed assemblies:")
        for name, err in failed:
            print(f"  {name}: {err}")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
