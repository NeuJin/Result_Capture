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

    # Use forward slashes — accepted by both PowerShell and Excel COM,
    # avoids all backslash-escaping issues in PowerShell double-quoted strings.
    def ps_path(p: Path) -> str:
        return str(p).replace("\\", "/")

    import_lines = []
    for mod in VBA_MODULES:
        import_lines.append(
            f'    $wb.VBProject.VBComponents.Import("{ps_path(VBA_DIR / mod)}") | Out-Null'
        )
    imports_block = "\n".join(import_lines)

    xlsm_path = ps_path(OUTPUT_XLSM)

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

# --- Close any leftover Excel COM instances from previous runs ---
try {{
    $stale = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Host "INFO: Found existing Excel instance — closing it first"
    $stale.DisplayAlerts = $false
    $stale.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($stale) | Out-Null
    Start-Sleep -Seconds 2
}} catch {{ }}  # no existing instance = fine

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {{
    $wb = $xl.Workbooks.Add(1)   # 1 = xlWBATWorksheet (normal blank workbook)

    # --- STEP 1: shape the workbook (sheets only, no VBA yet) ---
    while ($wb.Sheets.Count -lt 6) {{
        $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets($wb.Sheets.Count)) | Out-Null
    }}
    while ($wb.Sheets.Count -gt 6) {{
        $xl.DisplayAlerts = $false
        $wb.Sheets($wb.Sheets.Count).Delete()
    }}
{sheet_setup}

    # --- STEP 2: SaveAs xlsm NOW, while workbook is clean ---
    $xl.DisplayAlerts = $false
    $savePath = "{xlsm_path}"
    Write-Host "DEBUG: Saving to $savePath"
    Write-Host "DEBUG: Workbook name is $($wb.Name)"
    if (Test-Path $savePath) {{ Remove-Item $savePath -Force }}
    $wb.SaveAs($savePath, 52)
    Write-Host "SHELL SAVED: $savePath"

    # --- STEP 3: import VBA modules into the saved xlsm ---
{imports_block}

    # --- STEP 4: run setup macro + final save ---
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $xl.Run("SetupWorkbook.InitialiseWorkbook")
    Start-Sleep -Seconds 2
    $xl.DisplayAlerts = $false
    $wb.Save()

    Write-Host "DONE"
}}
catch {{
    Write-Host "ERROR: $_"
    exit 1
}}
finally {{
    # Keep Excel open so user can inspect the result
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
