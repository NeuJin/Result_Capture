Attribute VB_Name = "HVRunner"
Option Explicit

' ============================================================
' HVRunner — Write generated TCL to %TEMP%, launch hw.exe,
' poll for manifest file completion.
' ============================================================

Public Sub RunAnalysis()
    ' 1. Read & validate config
    Dim cfg As PipelineCfg
    cfg = ReadConfig()
    If Not ValidateConfig(cfg) Then Exit Sub

    ' 2. Ensure output folders exist
    EnsureOutputFolders cfg

    ' 3. Delete old manifest so WaitForManifest doesn't see stale file
    Dim manifestPath As String
    manifestPath = Replace(cfg.OutputFolder, "/", "\") & "\screenshot_manifest.txt"
    If FileExists(manifestPath) Then Kill manifestPath

    ' 4. Determine which TCL scripts to run
    Dim runStress As Boolean : runStress = (cfg.AnalysisType = "Stress" Or cfg.AnalysisType = "Both")
    Dim runSF     As Boolean : runSF     = (cfg.AnalysisType = "SafetyFactor" Or cfg.AnalysisType = "Both")

    ' 5. Run Stress analysis
    If runStress Then
        LogMessage "Generating Stress TCL script..."
        Dim stressTcl As String : stressTcl = GetStressTCL(cfg)
        Dim stressTclPath As String
        stressTclPath = Environ("TEMP") & "\conrod_stress.tcl"
        WriteTempTCL stressTcl, stressTclPath

        LogMessage "Launching HyperView (Stress)..."
        LaunchHV cfg.HVExePath, stressTclPath, cfg.LaunchMode

        LogMessage "Waiting for HyperView to complete..."
        If Not WaitForManifest(manifestPath, 600) Then
            MsgBox "Timeout: HyperView did not complete in 10 minutes (Stress)." & vbLf & _
                   "Check Log sheet for details.", vbCritical
            Exit Sub
        End If
        LogMessage "Stress analysis complete."
    End If

    ' 6. Run Safety Factor analysis (manifest appended or new)
    If runSF Then
        If runStress Then Kill manifestPath  ' reset manifest for SF run
        LogMessage "Generating Safety Factor TCL script..."
        Dim sfTcl As String : sfTcl = GetSafetyFactorTCL(cfg)
        Dim sfTclPath As String
        sfTclPath = Environ("TEMP") & "\conrod_sf.tcl"
        WriteTempTCL sfTcl, sfTclPath

        LogMessage "Launching HyperView (Safety Factor)..."
        LaunchHV cfg.HVExePath, sfTclPath, cfg.LaunchMode

        LogMessage "Waiting for HyperView to complete..."
        If Not WaitForManifest(manifestPath, 600) Then
            MsgBox "Timeout: HyperView did not complete in 10 minutes (Safety Factor).", vbCritical
            Exit Sub
        End If
        LogMessage "Safety Factor analysis complete."
    End If

    ' 7. Auto-fill Report sheet from manifest
    LogMessage "Filling Report sheet from screenshots..."
    FillReportFromManifest manifestPath

    LogMessage "DONE. Output folder: " & Replace(cfg.OutputFolder, "/", "\")
    MsgBox "Analysis complete!" & vbLf & vbLf & _
           "Screenshots and data saved to:" & vbLf & Replace(cfg.OutputFolder, "/", "\") & vbLf & vbLf & _
           "Review images in the Report sheet, then click [Generate PPTX].", _
           vbInformation, "Conrod Analysis Done"
End Sub

' ============================================================
' WriteTempTCL — write TCL string to file
' ============================================================
Private Sub WriteTempTCL(tclContent As String, filePath As String)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fso.CreateTextFile(filePath, True, False)  ' overwrite, not unicode
    f.Write tclContent
    f.Close
End Sub

' ============================================================
' LaunchHV — Shell hw.exe with TCL script
' ============================================================
Private Sub LaunchHV(hvExe As String, tclPath As String, launchMode As String)
    Dim cmd As String
    Dim quotedExe As String : quotedExe = """" & hvExe & """"
    Dim quotedTcl As String : quotedTcl = """" & tclPath & """"

    If UCase(launchMode) = "BATCH" Then
        cmd = quotedExe & " -b -tcl " & quotedTcl
    Else
        ' GUI mode: HV opens visible, sources TCL after load
        cmd = quotedExe & " -tcl " & quotedTcl
    End If

    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 1, False  ' window style 1 (normal), bWaitOnReturn=False (async)
End Sub

' ============================================================
' WaitForManifest — poll until manifest file exists (or timeout)
' Updates Log sheet every 3 seconds while waiting
' ============================================================
Public Function WaitForManifest(manifestPath As String, timeoutSec As Long) As Boolean
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim startTime As Single : startTime = Timer
    Dim elapsed As Long : elapsed = 0

    Do
        DoEvents
        Application.Wait Now + TimeValue("00:00:03")
        elapsed = CLng(Timer - startTime)
        LogMessage "Waiting for HyperView... " & elapsed & "s"

        If fso.FileExists(manifestPath) Then
            ' Wait a moment for file to finish writing
            Application.Wait Now + TimeValue("00:00:02")
            WaitForManifest = True
            Exit Function
        End If

        If elapsed > timeoutSec Then
            WaitForManifest = False
            Exit Function
        End If
    Loop
End Function

' ============================================================
' LogMessage — append timestamped line to Log sheet
' ============================================================
Public Sub LogMessage(msg As String)
    Dim wsh As Worksheet
    On Error Resume Next
    Set wsh = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    If wsh Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = wsh.Cells(wsh.Rows.Count, 1).End(xlUp).Row + 1
    wsh.Cells(nextRow, 1).Value = Now()
    wsh.Cells(nextRow, 2).Value = msg
    wsh.Cells(nextRow, 1).NumberFormat = "hh:mm:ss"

    ' Scroll to latest row
    Application.Goto wsh.Cells(nextRow, 1), True
    DoEvents
End Sub

' ============================================================
' ClearLog — clear Log sheet content (keep header)
' ============================================================
Public Sub ClearLog()
    Dim wsh As Worksheet : Set wsh = ThisWorkbook.Sheets("Log")
    Dim lastRow As Long : lastRow = wsh.Cells(wsh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then wsh.Rows("2:" & lastRow).ClearContents
End Sub
