"""
build_xlsm.py — Creates ConrodAnalysis.xlsm with all sheets + VBA modules.

Uses ONLY Python 3.8 standard library (subprocess, os, sys, tempfile, pathlib).
PowerShell handles Excel COM automation — no pip install required.

Run once:
    python build_xlsm.py

IMPORTANT: Excel must have VBA project trust enabled (one-time):
    Excel -> File -> Options -> Trust Center -> Trust Center Settings
    -> Macro Settings -> tick "Trust access to the VBA project object model"
"""

import os
import sys
import tempfile
import subprocess
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent.resolve()
VBA_DIR     = SCRIPT_DIR / "vba"
OUTPUT_XLSM = SCRIPT_DIR / "ConrodAnalysis.xlsm"

VBA_MODULES = [
    "ConfigManager.bas",
    "InpParser.bas",
    "TCLTemplates.bas",
    "HVRunner.bas",
    "ReportStager.bas",
    "PPTXBuilder.bas",
    "SetupWorkbook.bas",
]


def check_vba_files():
    missing = [m for m in VBA_MODULES if not (VBA_DIR / m).exists()]
    if missing:
        sys.exit(f"Missing VBA files in {VBA_DIR}:\n" + "\n".join(missing))


def build_powershell_script() -> str:
    """Generate the PowerShell script that creates the .xlsm via COM."""

    # Build the list of Import calls for each .bas file
    import_lines = []
    for mod in VBA_MODULES:
        path = str(VBA_DIR / mod).replace("\\", "\\\\")
        import_lines.append(f'    $wb.VBProject.VBComponents.Import("{path}") | Out-Null')
    imports_block = "\n".join(import_lines)

    xlsm_path = str(OUTPUT_XLSM).replace("\\", "\\\\")

    sheet_names = ["Config", "NodeSets", "Legend", "Views", "Report", "Log"]
    sheet_setup = ""
    for i, name in enumerate(sheet_names, start=1):
        sheet_setup += f"""
    if ($wb.Sheets.Count -lt {i}) {{
        $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count)) | Out-Null
    }}
    $wb.Sheets({i}).Name = "{name}"
"""

    script = f"""
$ErrorActionPreference = "Stop"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {{
    $wb = $xl.Workbooks.Add()

    # Ensure exactly 6 sheets, named correctly
    # Add sheets if fewer than needed
    while ($wb.Sheets.Count -lt 6) {{
        $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count)) | Out-Null
    }}
    # Remove extra sheets
    while ($wb.Sheets.Count -gt 6) {{
        $wb.Sheets($wb.Sheets.Count).Delete()
    }}

    # Rename sheets
{sheet_setup}

    # Import VBA modules
{imports_block}

    # Save as xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    if (Test-Path "{xlsm_path}") {{ Remove-Item "{xlsm_path}" -Force }}
    $wb.SaveAs("{xlsm_path}", 52)

    Write-Host "SAVED: {xlsm_path}"

    # Run InitialiseWorkbook to format all sheets
    $xl.Visible = $true
    $xl.Run("SetupWorkbook.InitialiseWorkbook")
    Start-Sleep -Seconds 2
    $wb.Save()

    Write-Host "DONE"
}}
catch {{
    Write-Host "ERROR: $_"
    exit 1
}}
finally {{
    # Keep Excel open so user can see result
}}
"""
    return script


def run():
    check_vba_files()

    print(f"Building: {OUTPUT_XLSM}")
    print(f"VBA dir : {VBA_DIR}")

    ps_script = build_powershell_script()

    # Write PS script to temp file
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", delete=False, encoding="utf-8"
    )
    tmp.write(ps_script)
    tmp.close()

    print("Running PowerShell Excel automation ...")
    try:
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", tmp.name,
            ],
            capture_output=True,
            text=True,
        )
        print(result.stdout.strip())
        if result.returncode != 0:
            print("STDERR:", result.stderr.strip())
            sys.exit(f"PowerShell exited with code {result.returncode}")
    finally:
        os.unlink(tmp.name)

    if OUTPUT_XLSM.exists():
        print(f"\nSuccess! File created: {OUTPUT_XLSM}")
        print("\nNext steps:")
        print("  1. Fill in Config sheet (result file, .inp, HyperView exe path)")
        print("  2. Click [Parse .inp] to load node sets")
        print("  3. Set Legend values, check Views")
        print("  4. Click [Run Analysis]")
    else:
        sys.exit("Build failed — xlsm file not created. Check Excel trust setting.")


if __name__ == "__main__":
    run()
