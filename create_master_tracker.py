"""
Create the master_tracker.xlsx file.
Shows each assembly # as rows, building block deliverables as columns.
Cells are populated with hyperlinks to deliverable files as they are generated.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

TRACKER_PATH = os.path.join(os.path.dirname(__file__), "assemblies", "master_tracker.xlsx")
ASSEMBLIES_DIR = os.path.join(os.path.dirname(__file__), "assemblies")

# Building block columns in progressive order
BUILDING_BLOCKS = [
    ("Source File", "source"),
    ("Cuttlet Plots", "cuttlet_plots"),
    ("Cuttlet Data", "cuttlet_data"),
    ("Cuttlet Forces", "cuttlet_forces"),
    ("Bit Body Forces", "bit_body_forces"),
    ("Min Engagement", "min_engagement"),
]


def find_assemblies():
    """Find all Assy_* folders in the assemblies directory."""
    assemblies = []
    if os.path.exists(ASSEMBLIES_DIR):
        for name in sorted(os.listdir(ASSEMBLIES_DIR)):
            if name.startswith("Assy_") and os.path.isdir(os.path.join(ASSEMBLIES_DIR, name)):
                assy_num = name.replace("Assy_", "")
                assemblies.append((assy_num, name))
    return assemblies


def find_deliverable(assy_folder, bb_subfolder):
    """Find the first file in a building block subfolder. Returns relative path or None."""
    bb_path = os.path.join(ASSEMBLIES_DIR, assy_folder, bb_subfolder)
    if os.path.exists(bb_path):
        files = [f for f in os.listdir(bb_path) if not f.startswith(".")]
        if files:
            return os.path.join(assy_folder, bb_subfolder, files[0])
    return None


def create_tracker():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Master Tracker"

    # Styles
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    link_font = Font(color="0563C1", underline="single", size=11)
    pending_font = Font(color="999999", italic=True, size=11)

    # Title row
    ws.merge_cells("A1:G1")
    ws["A1"] = "Bit Design Building Blocks — Master Tracker"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")

    # Header row
    headers = ["Assy #"] + [bb[0] for bb in BUILDING_BLOCKS]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Set column widths
    ws.column_dimensions["A"].width = 14
    for col in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20

    # Row height for header
    ws.row_dimensions[3].height = 30

    # Populate assembly rows
    assemblies = find_assemblies()
    for row_idx, (assy_num, assy_folder) in enumerate(assemblies, 4):
        # Assy # column
        cell = ws.cell(row=row_idx, column=1, value=assy_num)
        cell.font = Font(bold=True, size=11)
        cell.alignment = cell_align
        cell.border = thin_border

        # Building block columns
        for col_idx, (bb_name, bb_subfolder) in enumerate(BUILDING_BLOCKS, 2):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = cell_align

            deliverable = find_deliverable(assy_folder, bb_subfolder)
            if deliverable:
                # Hyperlink to the file (relative path from tracker location)
                cell.value = bb_name
                cell.hyperlink = deliverable
                cell.font = link_font
            else:
                cell.value = "—"
                cell.font = pending_font

    wb.save(TRACKER_PATH)
    print(f"Master tracker saved to: {TRACKER_PATH}")
    print(f"  Assemblies found: {len(assemblies)}")
    for assy_num, assy_folder in assemblies:
        linked = sum(1 for _, bb_sub in BUILDING_BLOCKS if find_deliverable(assy_folder, bb_sub))
        print(f"    {assy_num}: {linked}/{len(BUILDING_BLOCKS)} deliverables linked")


if __name__ == "__main__":
    create_tracker()
