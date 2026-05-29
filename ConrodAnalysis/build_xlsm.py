"""
build_xlsm.py — Creates ConrodAnalysis.xlsm with all sheets formatted
and all VBA modules imported.

Requirements:
  pip install pywin32

Run once:
  python build_xlsm.py

IMPORTANT: Excel must have "Trust access to the VBA project object model"
  enabled: File → Options → Trust Center → Trust Center Settings
           → Macro Settings → tick "Trust access to the VBA project object model"
"""

import os
import sys
import time

try:
    import win32com.client as win32
except ImportError:
    sys.exit("pywin32 not found. Run: pip install pywin32")

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
VBA_DIR      = os.path.join(SCRIPT_DIR, "vba")
OUTPUT_XLSM  = os.path.join(SCRIPT_DIR, "ConrodAnalysis.xlsm")

VBA_MODULES = [
    "ConfigManager.bas",
    "InpParser.bas",
    "TCLTemplates.bas",
    "HVRunner.bas",
    "ReportStager.bas",
    "PPTXBuilder.bas",
    "SetupWorkbook.bas",
]


def main():
    print("Building ConrodAnalysis.xlsm ...")

    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    # New workbook
    wb = xl.Workbooks.Add()

    # Pre-create sheets so SetupWorkbook.InitialiseWorkbook can find them
    sheet_names = ["Config", "NodeSets", "Legend", "Views", "Report", "Log"]
    # Rename existing sheets first
    existing = [wb.Sheets(i + 1).Name for i in range(wb.Sheets.Count)]
    for i, name in enumerate(sheet_names):
        if i < len(existing):
            wb.Sheets(i + 1).Name = name
        else:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count)).Name = name

    # Remove extra default sheets
    while wb.Sheets.Count > len(sheet_names):
        wb.Sheets(wb.Sheets.Count).Delete

    # Import VBA modules
    vba_project = wb.VBProject
    for mod_file in VBA_MODULES:
        mod_path = os.path.join(VBA_DIR, mod_file)
        if not os.path.exists(mod_path):
            print(f"  WARNING: {mod_file} not found, skipping")
            continue
        vba_project.VBComponents.Import(mod_path)
        print(f"  Imported: {mod_file}")

    # Save as xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    if os.path.exists(OUTPUT_XLSM):
        os.remove(OUTPUT_XLSM)
    wb.SaveAs(OUTPUT_XLSM, FileFormat=52)
    print(f"  Saved: {OUTPUT_XLSM}")

    # Run InitialiseWorkbook to set up all sheets
    print("  Running InitialiseWorkbook ...")
    xl.Visible = True
    time.sleep(1)
    xl.Run("SetupWorkbook.InitialiseWorkbook")
    time.sleep(2)

    wb.Save()
    print("\nDone! ConrodAnalysis.xlsm is ready.")
    print("Next steps:")
    print("  1. Fill in Config sheet (result file, .inp path, HyperView exe)")
    print("  2. Click [Parse .inp] to populate NodeSets")
    print("  3. Set Legend values, check Views")
    print("  4. Click [Run Analysis]")


if __name__ == "__main__":
    main()
