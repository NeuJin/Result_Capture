Attribute VB_Name = "ConfigManager"
Option Explicit

' ============================================================
' ConfigManager — Read all sheet inputs into a config object,
' validate, and expose helpers for other modules.
' ============================================================

' Simple config container (use Public Type for easy passing between modules)
Public Type LegendCfg
    MinVal      As Double
    MaxVal      As Double
    NumLevels   As Integer
    ScaleType   As String   ' "Linear" or "Log"
    Palette     As String   ' "Rainbow" / "Thermal" / "Blue-Red"
    ShowValues  As Boolean
    Discrete    As Boolean  ' True=Discrete, False=Continuous
End Type

Public Type ViewsCfg
    AutoFit     As Boolean
    ShowISO     As Boolean
    ShowFront   As Boolean
    ShowBack    As Boolean
    ShowLeft    As Boolean
    ShowRight   As Boolean
    ShowTop     As Boolean
    ScrW        As Long
    ScrH        As Long
End Type

Public Type NodeSetEntry
    SetID       As Long
    NsetName    As String
    Label       As String
    NodeIDs     As String   ' comma-separated (from Col F)
End Type

Public Type PipelineCfg
    ResultFile   As String
    InpFile      As String
    AnalysisType As String  ' "Stress" / "SafetyFactor" / "Both"
    HVExePath    As String
    OutputFolder As String
    LaunchMode   As String  ' "Batch" / "GUI"
    OutputFormat As String  ' "Images" / "PPTX" / "Both"
    Leg          As LegendCfg
    Views        As ViewsCfg
    NodeSets()   As NodeSetEntry
    NodeSetCount As Integer
End Type

' ============================================================
' ReadConfig — build PipelineCfg from all sheets
' ============================================================
Public Function ReadConfig() As PipelineCfg
    Dim cfg As PipelineCfg
    Dim cfgSh As Worksheet : Set cfgSh = ThisWorkbook.Sheets("Config")
    Dim legSh As Worksheet : Set legSh = ThisWorkbook.Sheets("Legend")
    Dim vwSh  As Worksheet : Set vwSh  = ThisWorkbook.Sheets("Views")
    Dim nsSh  As Worksheet : Set nsSh  = ThisWorkbook.Sheets("NodeSets")

    ' --- Config sheet ---
    cfg.ResultFile   = Trim(cfgSh.Range("B3").Value)
    cfg.InpFile      = Trim(cfgSh.Range("B4").Value)
    cfg.AnalysisType = Trim(cfgSh.Range("B5").Value)
    cfg.HVExePath    = Trim(cfgSh.Range("B6").Value)
    cfg.OutputFolder = Trim(cfgSh.Range("B7").Value)
    cfg.LaunchMode   = Trim(cfgSh.Range("B8").Value)
    cfg.OutputFormat = Trim(cfgSh.Range("B9").Value)

    ' Normalise output folder — no trailing backslash, forward slashes for TCL
    cfg.OutputFolder = Replace(cfg.OutputFolder, "\", "/")
    If Right(cfg.OutputFolder, 1) = "/" Then
        cfg.OutputFolder = Left(cfg.OutputFolder, Len(cfg.OutputFolder) - 1)
    End If

    ' --- Legend sheet ---
    With cfg.Leg
        .MinVal     = Val(legSh.Range("B3").Value)
        .MaxVal     = Val(legSh.Range("B4").Value)
        .NumLevels  = CInt(Val(legSh.Range("B5").Value))
        .ScaleType  = Trim(legSh.Range("B6").Value)
        .Palette    = Trim(legSh.Range("B7").Value)
        .ShowValues = (UCase(Trim(legSh.Range("B8").Value)) = "Y")
        .Discrete   = (UCase(Trim(legSh.Range("B9").Value)) = "DISCRETE")
    End With

    ' --- Views sheet ---
    With cfg.Views
        .AutoFit   = (UCase(Trim(vwSh.Range("B3").Value)) = "Y")
        .ShowISO   = (UCase(Trim(vwSh.Range("B5").Value)) = "Y")
        .ShowFront = (UCase(Trim(vwSh.Range("B6").Value)) = "Y")
        .ShowBack  = (UCase(Trim(vwSh.Range("B7").Value)) = "Y")
        .ShowLeft  = (UCase(Trim(vwSh.Range("B8").Value)) = "Y")
        .ShowRight = (UCase(Trim(vwSh.Range("B9").Value)) = "Y")
        .ShowTop   = (UCase(Trim(vwSh.Range("B10").Value)) = "Y")
        .ScrW      = CLng(Val(vwSh.Range("B12").Value))
        .ScrH      = CLng(Val(vwSh.Range("B13").Value))
        If .ScrW <= 0 Then .ScrW = 1920
        If .ScrH <= 0 Then .ScrH = 1080
    End With

    ' --- NodeSets sheet (only Include=Y rows) ---
    Dim lastRow As Long
    lastRow = nsSh.Cells(nsSh.Rows.Count, 1).End(xlUp).Row
    Dim count As Integer : count = 0

    ReDim cfg.NodeSets(0 To 100)  ' max 100 sets; resize below

    Dim r As Long
    For r = 2 To lastRow
        If UCase(Trim(nsSh.Cells(r, 3).Value)) = "Y" Then
            cfg.NodeSets(count).SetID    = CLng(nsSh.Cells(r, 1).Value)
            cfg.NodeSets(count).NsetName = Trim(nsSh.Cells(r, 2).Value)
            cfg.NodeSets(count).Label    = Trim(nsSh.Cells(r, 4).Value)
            cfg.NodeSets(count).NodeIDs  = Trim(nsSh.Cells(r, 6).Value)
            count = count + 1
        End If
    Next r

    cfg.NodeSetCount = count
    ReDim Preserve cfg.NodeSets(0 To count - 1)

    ReadConfig = cfg
End Function

' ============================================================
' ValidateConfig — returns True if OK, False + MsgBox if not
' ============================================================
Public Function ValidateConfig(cfg As PipelineCfg) As Boolean
    Dim errors As String : errors = ""
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")

    If cfg.ResultFile = "" Then errors = errors & "- Result file path is empty" & vbLf
    If cfg.ResultFile <> "" And Not fso.FileExists(cfg.ResultFile) Then
        errors = errors & "- Result file not found: " & cfg.ResultFile & vbLf
    End If

    If cfg.HVExePath = "" Then errors = errors & "- HyperView exe path is empty" & vbLf
    If cfg.HVExePath <> "" And Not fso.FileExists(cfg.HVExePath) Then
        errors = errors & "- HyperView exe not found: " & cfg.HVExePath & vbLf
    End If

    If cfg.OutputFolder = "" Then errors = errors & "- Output folder is empty" & vbLf

    If cfg.NodeSetCount = 0 Then
        errors = errors & "- No node sets selected (check NodeSets sheet, Include column)" & vbLf
    End If

    If cfg.Leg.MaxVal <= cfg.Leg.MinVal Then
        errors = errors & "- Legend Max must be greater than Min" & vbLf
    End If

    If Not (cfg.Views.ShowISO Or cfg.Views.ShowFront Or cfg.Views.ShowBack Or _
            cfg.Views.ShowLeft Or cfg.Views.ShowRight Or cfg.Views.ShowTop) Then
        errors = errors & "- No views selected (check Views sheet)" & vbLf
    End If

    If errors <> "" Then
        MsgBox "Please fix the following before running:" & vbLf & vbLf & errors, vbCritical, "Validation Failed"
        ValidateConfig = False
    Else
        ValidateConfig = True
    End If
End Function

' ============================================================
' EnsureOutputFolders — create output subfolders if missing
' ============================================================
Public Sub EnsureOutputFolders(cfg As PipelineCfg)
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String : base = Replace(cfg.OutputFolder, "/", "\")

    Dim subFolders(2) As String
    subFolders(0) = base & "\screenshots"
    subFolders(1) = base & "\csv"
    subFolders(2) = base & "\reports"

    Dim i As Integer
    For i = 0 To 2
        If Not fso.FolderExists(subFolders(i)) Then
            fso.CreateFolder subFolders(i)
        End If
    Next i
End Sub

' ============================================================
' BoolToTCL — convert VBA Boolean to TCL "1"/"0"
' ============================================================
Public Function BoolToTCL(b As Boolean) As String
    BoolToTCL = IIf(b, "1", "0")
End Function

' ============================================================
' PathToTCL — ensure forward slashes for TCL compatibility
' ============================================================
Public Function PathToTCL(p As String) As String
    PathToTCL = Replace(p, "\", "/")
End Function
