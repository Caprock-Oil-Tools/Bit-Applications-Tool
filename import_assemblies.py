"""
Import all assemblies into the Master folder structure.

For each Bit Designs folder:
1. Find the highest-version Min Engagement source file (skip IPR variants, temp files)
2. Extract the assembly name from the filename (everything before " Min Engagement")
3. Copy the source file to Master/<assy_name>/
4. Rebuild the Master List
"""

import os
import re
import shutil

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
BIT_DESIGNS_DIR = os.path.join(PROJECT_ROOT, "Bit Designs")
MASTER_DIR = os.path.join(PROJECT_ROOT, "Master")


def parse_version(filename):
    """Extract version number from filename like 'v6.07' -> (6, 7), 'v8.01' -> (8, 1)."""
    m = re.search(r'v(\d+)\.(\d+)', filename)
    if m:
        return (int(m.group(1)), int(m.group(2)))
    return (0, 0)


def extract_assy_name(filename):
    """Extract assembly name from filename. E.g. '2176r06 Min Engagement v8.01.xlsm' -> '2176r06'."""
    # Remove everything from " Min Engagement" onward
    idx = filename.find(" Min Engagement")
    if idx > 0:
        return filename[:idx]
    return None


def is_ipr_variant(filename):
    """Check if this is an IPR-specific variant like '@ 0.100 in_rev'."""
    return "@ " in filename or "in_rev" in filename or "0.100" in filename or "0.500" in filename


def is_temp_file(filename):
    """Check if this is a temp/test file."""
    return filename.startswith("~$") or filename.startswith("Test_")


def find_best_source_files():
    """Find the best (highest version) source file for each assembly folder."""
    results = []

    for folder_name in sorted(os.listdir(BIT_DESIGNS_DIR)):
        folder_path = os.path.join(BIT_DESIGNS_DIR, folder_name)
        if not os.path.isdir(folder_path):
            continue

        # Find all Min Engagement files in this folder (not in subfolders for now)
        candidates = []
        for fname in os.listdir(folder_path):
            if "Min Engagement" in fname and fname.endswith(".xlsm"):
                if is_temp_file(fname) or is_ipr_variant(fname):
                    continue
                assy_name = extract_assy_name(fname)
                if assy_name:
                    version = parse_version(fname)
                    candidates.append((assy_name, fname, version))

        # Also check one level of subfolders (e.g., "2265 w Rippers/")
        for sub in os.listdir(folder_path):
            sub_path = os.path.join(folder_path, sub)
            if os.path.isdir(sub_path):
                for fname in os.listdir(sub_path):
                    if "Min Engagement" in fname and fname.endswith(".xlsm"):
                        if is_temp_file(fname) or is_ipr_variant(fname):
                            continue
                        assy_name = extract_assy_name(fname)
                        if assy_name:
                            version = parse_version(fname)
                            candidates.append((assy_name, os.path.join(sub, fname), version))

        if not candidates:
            continue

        # Group by assy_name, pick highest version for each
        by_name = {}
        for assy_name, fname, version in candidates:
            if assy_name not in by_name or version > by_name[assy_name][1]:
                by_name[assy_name] = (fname, version)

        for assy_name, (fname, version) in sorted(by_name.items()):
            full_path = os.path.join(folder_path, fname)
            results.append((folder_name, assy_name, fname, full_path))

    return results


def import_all():
    os.makedirs(MASTER_DIR, exist_ok=True)

    sources = find_best_source_files()
    print(f"Found {len(sources)} assemblies with source files.\n")

    imported = 0
    skipped = 0

    for folder_name, assy_name, fname, full_path in sources:
        # Normalize assy_name: remove dots for consistency (e.g. "2004.r1" -> "2004r1")
        # But keep the original filename as-is
        assy_folder = assy_name.replace(".", "")
        dest_dir = os.path.join(MASTER_DIR, assy_folder)
        dest_file = os.path.join(dest_dir, os.path.basename(fname))

        if os.path.exists(dest_file):
            skipped += 1
            continue

        os.makedirs(dest_dir, exist_ok=True)
        shutil.copy2(full_path, dest_file)
        imported += 1
        print(f"  {assy_folder}: {os.path.basename(fname)}")

    print(f"\nImported: {imported}, Already existed: {skipped}")
    print(f"Total assemblies in Master: {len([d for d in os.listdir(MASTER_DIR) if os.path.isdir(os.path.join(MASTER_DIR, d))])}")


if __name__ == "__main__":
    import_all()
