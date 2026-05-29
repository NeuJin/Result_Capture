Attribute VB_Name = "SetupWorkbook"
Option Explicit

' ============================================================
' SetupWorkbook — initialize all sheets with headers, labels,
' dropdowns, formatting, and buttons.
' Run ONCE after importing all VBA modules.
' ============================================================

Public Sub InitialiseWorkbook()
    Application.ScreenUpdating = False

    SetupConfigSheet
    SetupNodeSetsSheet
    SetupLegendSheet
    SetupViewsSheet
    SetupReportSheet  ' from ReportStager module
    SetupLogSheet

    ' Set tab order and colors
    ThisWorkbook.Sheets("Config").Tab.Color   = RGB(0, 112, 192)
    ThisWorkbook.Sheets("NodeSets").Tab.Color  = RGB(0, 176, 80)
    ThisWorkbook.Sheets("Legend").Tab.Color    = RGB(255, 192, 0)
    ThisWorkbook.Sheets("Views").Tab.Color     = RGB(146, 208, 80)
    ThisWorkbook.Sheets("Report").Tab.Color    = RGB(255, 0, 0)
    ThisWorkbook.Sheets("Log").Tab.Color       = RGB(128, 128, 128)

    ThisWorkbook.Sheets("Config").Activate
    Application.ScreenUpdating = True
    MsgBox "Workbook initialised! Fill in Config sheet, then click [Parse .inp] to start.", _
           vbInformation, "Setup Complete"
End Sub

' ============================================================
' Config Sheet
' ============================================================
Private Sub SetupConfigSheet()
    Dim wsh As Worksheet : Set wsh = GetOrCreateSheet("Config")
    wsh.Cells.ClearContents : wsh.Cells.ClearFormats

    ' Title
    FormatTitle wsh, "CONROD ANALYSIS — CONFIGURATION"

    Dim labels() As Variant
    labels = Array( _
        "Result File (.odb / .res / .h3d / .op2):", _
        "Model File (.inp — Abaqus):", _
        "Analysis Type:", _
        "HyperView Exe Path:", _
        "Output Folder:", _
        "Launch Mode:", _
        "Output Format:", _
        "Screenshot Resolution (W x H):", _
        "", _
        "Python Exe (optional):", _
        "Use Python Backend:" _
    )
    Dim i As Integer
    For i = 0 To UBound(labels)
        wsh.Cells(3 + i, 1).Value = labels(i)
        wsh.Cells(3 + i, 1).Font.Bold = True
        wsh.Cells(3 + i, 1).HorizontalAlignment = xlRight
    Next i

    ' Default values in Col B
    wsh.Range("B5").Value = "Stress"       ' Analysis Type default
    wsh.Range("B8").Value = "Batch"        ' Launch Mode default
    wsh.Range("B9").Value = "PPTX"         ' Output Format default
    wsh.Range("B10").Value = "1920 x 1080" ' Resolution default
    wsh.Range("B13").Value = "N"           ' Use Python Backend default

    ' Dropdown validation
    AddDropdown wsh, "B5", "Stress,SafetyFactor,Both"
    AddDropdown wsh, "B8", "Batch,GUI"
    AddDropdown wsh, "B9", "Images,PPTX,Both"
    AddDropdown wsh, "B13", "Y,N"

    ' Col widths
    wsh.Columns("A").ColumnWidth = 40
    wsh.Columns("B").ColumnWidth = 60

    ' Buttons
    AddButton wsh, "Parse .inp → NodeSets", "InpParser.ParseInpFile", 250, 14, 200, 26
    AddButton wsh, "Run Analysis",           "HVRunner.RunAnalysis",  460, 14, 150, 26
    AddButton wsh, "Open Output Folder",     "OpenOutputFolder",      620, 14, 140, 26
End Sub

' ============================================================
' NodeSets Sheet
' ============================================================
Private Sub SetupNodeSetsSheet()
    Dim wsh As Worksheet : Set wsh = GetOrCreateSheet("NodeSets")
    wsh.Cells.ClearContents : wsh.Cells.ClearFormats

    FormatTitle wsh, "NODE SETS — from .inp model"

    Dim headers() As Variant
    headers = Array("Set ID", "NSET Name", "Include?", "Display Label", "Node Count", "Node IDs (auto)")
    Dim c As Integer
    For c = 0 To UBound(headers)
        With wsh.Cells(3, c + 1)
            .Value = headers(c)
            .Font.Bold = True
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next c

    ' Column widths
    wsh.Columns("A").ColumnWidth = 8
    wsh.Columns("B").ColumnWidth = 25
    wsh.Columns("C").ColumnWidth = 10
    wsh.Columns("D").ColumnWidth = 25
    wsh.Columns("E").ColumnWidth = 12
    wsh.Columns("F").ColumnWidth = 60

    ' Col F header note
    wsh.Cells(4, 6).Value = "(auto-filled — do not edit)"
    wsh.Cells(4, 6).Font.Italic = True
    wsh.Cells(4, 6).Font.Color = RGB(150, 150, 150)

    ' Hide Col F (raw node IDs — for internal use)
    wsh.Columns("F").Hidden = True

    wsh.Rows(3).RowHeight = 18
End Sub

' ============================================================
' Legend Sheet
' ============================================================
Private Sub SetupLegendSheet()
    Dim wsh As Worksheet : Set wsh = GetOrCreateSheet("Legend")
    wsh.Cells.ClearContents : wsh.Cells.ClearFormats

    FormatTitle wsh, "LEGEND SETTINGS — Full User Control"

    wsh.Cells(2, 1).Value = "All values are set MANUALLY. The tool never auto-adjusts the legend."
    wsh.Cells(2, 1).Font.Italic = True
    wsh.Cells(2, 1).Font.Color = RGB(180, 0, 0)

    Dim rows() As Variant
    rows = Array( _
        Array("Min Value", 0, "Lower bound of legend (e.g. 0 for stress)"), _
        Array("Max Value", 500, "Upper bound — set to expected max before running"), _
        Array("Number of Levels", 10, "Colour bands (5–20 typical)"), _
        Array("Scale Type", "Linear", "Linear or Log"), _
        Array("Color Palette", "Rainbow", "Rainbow / Thermal / Blue-Red"), _
        Array("Show Values on Legend", "Y", "Y = show numeric labels on legend"), _
        Array("Style", "Discrete", "Discrete (banded) or Continuous (smooth)") _
    )

    Dim i As Integer
    For i = 0 To UBound(rows)
        With wsh.Cells(3 + i, 1)
            .Value = rows(i)(0)
            .Font.Bold = True
            .HorizontalAlignment = xlRight
        End With
        wsh.Cells(3 + i, 2).Value = rows(i)(1)
        wsh.Cells(3 + i, 3).Value = rows(i)(2)
        wsh.Cells(3 + i, 3).Font.Color = RGB(120, 120, 120)
        wsh.Cells(3 + i, 3).Font.Italic = True
    Next i

    ' Dropdowns
    AddDropdown wsh, "B6", "Linear,Log"
    AddDropdown wsh, "B7", "Rainbow,Thermal,Blue-Red"
    AddDropdown wsh, "B8", "Y,N"
    AddDropdown wsh, "B9", "Discrete,Continuous"

    wsh.Columns("A").ColumnWidth = 28
    wsh.Columns("B").ColumnWidth = 15
    wsh.Columns("C").ColumnWidth = 50

    ' Highlight input cells
    wsh.Range("B3:B9").Interior.Color = RGB(255, 255, 204)
End Sub

' ============================================================
' Views Sheet
' ============================================================
Private Sub SetupViewsSheet()
    Dim wsh As Worksheet : Set wsh = GetOrCreateSheet("Views")
    wsh.Cells.ClearContents : wsh.Cells.ClearFormats

    FormatTitle wsh, "VIEW SETTINGS — Screenshot Configuration"

    wsh.Cells(2, 1).Value = "Camera auto-fits to the node region (masked model) before each capture."
    wsh.Cells(2, 1).Font.Italic = True

    ' Auto-fit toggle
    wsh.Cells(3, 1).Value = "Auto-Fit to Node Region:" : wsh.Cells(3, 1).Font.Bold = True
    wsh.Cells(3, 2).Value = "Y"
    AddDropdown wsh, "B3", "Y,N"

    ' Separator
    wsh.Cells(4, 1).Value = "— Views to capture (Y = enabled) —"
    wsh.Cells(4, 1).Font.Bold = True
    wsh.Cells(4, 1).Font.Color = RGB(0, 70, 127)

    Dim viewRows() As Variant
    viewRows = Array( _
        Array("ISO (isometric)",   "Y"), _
        Array("Front (+Y direction)",  "Y"), _
        Array("Back (-Y direction)",   "Y"), _
        Array("Left (+X direction)",   "Y"), _
        Array("Right (-X direction)",  "Y"), _
        Array("Top (+Z direction)",    "Y") _
    )

    Dim i As Integer
    For i = 0 To UBound(viewRows)
        wsh.Cells(5 + i, 1).Value = viewRows(i)(0)
        wsh.Cells(5 + i, 1).Font.Bold = True
        wsh.Cells(5 + i, 2).Value = viewRows(i)(1)
        AddDropdown wsh, "B" & (5 + i), "Y,N"
    Next i

    ' Resolution
    wsh.Cells(12, 1).Value = "Screenshot Width (px):"  : wsh.Cells(12, 1).Font.Bold = True
    wsh.Cells(12, 2).Value = 1920
    wsh.Cells(13, 1).Value = "Screenshot Height (px):" : wsh.Cells(13, 1).Font.Bold = True
    wsh.Cells(13, 2).Value = 1080

    wsh.Columns("A").ColumnWidth = 28
    wsh.Columns("B").ColumnWidth = 10
    wsh.Range("B3:B13").Interior.Color = RGB(255, 255, 204)
End Sub

' ============================================================
' Log Sheet
' ============================================================
Private Sub SetupLogSheet()
    Dim wsh As Worksheet : Set wsh = GetOrCreateSheet("Log")
    wsh.Cells.ClearContents : wsh.Cells.ClearFormats

    FormatTitle wsh, "ANALYSIS LOG"

    With wsh.Cells(3, 1) : .Value = "Time"     : .Font.Bold = True : End With
    With wsh.Cells(3, 2) : .Value = "Message"  : .Font.Bold = True : End With

    wsh.Columns("A").ColumnWidth = 12
    wsh.Columns("B").ColumnWidth = 80

    AddButton wsh, "Clear Log", "HVRunner.ClearLog", 10, 10, 100, 22
End Sub

' ============================================================
' Helper: get or create a sheet by name
' ============================================================
Private Function GetOrCreateSheet(name As String) As Worksheet
    Dim wsh As Worksheet
    On Error Resume Next
    Set wsh = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If wsh Is Nothing Then
        Set wsh = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsh.Name = name
    End If
    Set GetOrCreateSheet = wsh
End Function

' ============================================================
' Helper: format title row
' ============================================================
Private Sub FormatTitle(wsh As Worksheet, titleText As String)
    With wsh.Range(wsh.Cells(1, 1), wsh.Cells(1, 8))
        .Merge
        .Value = titleText
        .Interior.Color = RGB(0, 70, 127)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
        .RowHeight = 26
    End With
End Sub

' ============================================================
' Helper: add data validation dropdown
' ============================================================
Private Sub AddDropdown(wsh As Worksheet, cellAddr As String, choices As String)
    With wsh.Range(cellAddr).Validation
        .Delete
        .Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:=choices  ' xlValidateList
        .ShowDropDown = False
    End With
End Sub

' ============================================================
' Helper: add a form button linked to a macro
' ============================================================
Private Sub AddButton(wsh As Worksheet, caption As String, macro As String, _
    leftPt As Single, topPt As Single, w As Single, h As Single)

    Dim btn As Object
    Set btn = wsh.Buttons.Add(leftPt, topPt, w, h)
    btn.Caption = caption
    btn.OnAction = macro
    btn.Font.Size = 9
End Sub

' ============================================================
' OpenOutputFolder — shell explorer to output path
' ============================================================
Public Sub OpenOutputFolder()
    Dim cfg As PipelineCfg : cfg = ReadConfig()
    Dim path As String : path = Replace(cfg.OutputFolder, "/", "\")
    If path <> "" Then Shell "explorer.exe " & Chr(34) & path & Chr(34), vbNormalFocus
End Sub
