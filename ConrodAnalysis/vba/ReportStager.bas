Attribute VB_Name = "ReportStager"
Option Explicit

' ============================================================
' ReportStager — Read screenshot_manifest.txt and place images
' into the Report sheet slide blocks.
'
' Manifest format (tab-separated):
'   set_id  TAB  label  TAB  view  TAB  image_path  TAB  value
'
' Report sheet layout per slide block (BLOCK_HEIGHT = 17 rows):
'   Row +0  : Title bar (merged A:F)
'   Row +1  : View labels row 1: ISO | Front | Back
'   Row +2..+7 : Image row 1 (each image spans 6 rows)
'   Row +8  : View labels row 2: Left | Right | Top
'   Row +9..+14: Image row 2
'   Row +15 : Footer value text (merged A:F)
'   Row +16 : Separator (thin border)
' ============================================================

Private Const BLOCK_HEIGHT  As Long = 17
Private Const BLOCK_IMG_H   As Long = 6   ' rows per image
Private Const DATA_START_ROW As Long = 3  ' first block starts at row 3

Private Const NAVY_COLOR    As Long = 3368601  ' RGB(51,102,153) dark blue
Private Const LIGHT_GRAY    As Long = 15921906 ' RGB(242,242,242)

' ============================================================
' SetupReportSheet — format Report sheet with a blank template
' block so the user can see the structure.
' Call once when building the workbook for the first time.
' ============================================================
Public Sub SetupReportSheet()
    Dim wsh As Worksheet : Set wsh = ThisWorkbook.Sheets("Report")
    wsh.Cells.ClearContents
    wsh.Cells.ClearFormats

    ' ---- Header row ----
    With wsh.Cells(1, 1)
        .Value = "CONROD ANALYSIS — REPORT STAGING"
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsh.Cells(2, 1).Value = "Auto-filled after analysis. Edit captions or swap images, then click [Generate PPTX]."

    ' ---- Buttons row ----
    ' (Buttons are added programmatically by SetupWorkbook; here just set row height)
    wsh.Rows(2).RowHeight = 20

    ' ---- Template (placeholder) block ----
    BuildTemplateBlock wsh, DATA_START_ROW, 0, "(Set Label)", "(Analysis Type)", "(Value)"

    ' Column widths for 3-column image layout
    wsh.Columns("A").ColumnWidth = 30
    wsh.Columns("B").ColumnWidth = 2   ' spacer
    wsh.Columns("C").ColumnWidth = 30
    wsh.Columns("D").ColumnWidth = 2   ' spacer
    wsh.Columns("E").ColumnWidth = 30
    wsh.Columns("F").ColumnWidth = 5   ' right margin

    LogMessage "Report sheet template created."
End Sub

' ============================================================
' FillReportFromManifest — read manifest, populate blocks
' ============================================================
Public Sub FillReportFromManifest(manifestPath As String)
    If Not FileExists(manifestPath) Then
        MsgBox "Manifest not found: " & manifestPath, vbCritical
        Exit Sub
    End If

    Dim wsh As Worksheet : Set wsh = ThisWorkbook.Sheets("Report")

    ' Parse manifest into groups: set_id -> {view -> path, label, value}
    Dim groups As Object : Set groups = CreateObject("Scripting.Dictionary")  ' set_id -> Object
    Dim groupOrder() As Long
    Dim groupCount As Long : groupCount = 0

    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object : Set f = fso.OpenTextFile(manifestPath, 1)

    ReDim groupOrder(0 To 200)

    Do While Not f.AtEndOfStream
        Dim line As String : line = Trim(f.ReadLine)
        If line = "" Then GoTo NextManifestLine

        Dim parts() As String : parts = Split(line, vbTab)
        If UBound(parts) < 4 Then GoTo NextManifestLine

        Dim setID  As Long   : setID  = CLng(Trim(parts(0)))
        Dim lbl    As String : lbl    = Trim(parts(1))
        Dim vname  As String : vname  = Trim(parts(2))
        Dim imgPth As String : imgPth = Replace(Trim(parts(3)), "/", "\")
        Dim val    As String : val    = Trim(parts(4))

        If Not groups.Exists(setID) Then
            Dim grp As Object : Set grp = CreateObject("Scripting.Dictionary")
            grp("label") = lbl
            grp("value") = val
            Set groups(setID) = grp
            groupOrder(groupCount) = setID
            groupCount = groupCount + 1
        End If

        Dim g As Object : Set g = groups(setID)
        g(vname) = imgPth  ' "ISO" -> "C:\...\ISO.png"

NextManifestLine:
    Loop
    f.Close

    If groupCount = 0 Then
        MsgBox "No data found in manifest file.", vbExclamation
        Exit Sub
    End If

    ' --- Clear existing images from Report sheet ---
    ClearAllImages wsh

    ' --- Clear old content below header ---
    Dim lastRow As Long : lastRow = wsh.Cells(wsh.Rows.Count, 1).End(xlUp).Row
    If lastRow >= DATA_START_ROW Then
        wsh.Rows(DATA_START_ROW & ":" & lastRow).ClearContents
        wsh.Rows(DATA_START_ROW & ":" & lastRow).ClearFormats
    End If

    ' --- Build blocks ---
    Dim blockIdx As Long
    For blockIdx = 0 To groupCount - 1
        Dim sid As Long : sid = groupOrder(blockIdx)
        Dim gd As Object : Set gd = groups(sid)

        Dim startRow As Long
        startRow = DATA_START_ROW + blockIdx * BLOCK_HEIGHT

        ' Determine analysis label for title
        Dim analysisLabel As String
        If gd.Exists("ISO") Or gd.Exists("Front") Then
            analysisLabel = "Von Mises Stress"
        Else
            analysisLabel = "Safety Factor"
        End If

        BuildTemplateBlock wsh, startRow, sid, gd("label"), analysisLabel, gd("value")

        ' --- Place images ---
        ' Row 1 image positions: cols A(1), C(3), E(5)
        Dim viewRow1(2) As String : viewRow1(0) = "ISO" : viewRow1(1) = "Front" : viewRow1(2) = "Back"
        Dim viewRow2(2) As String : viewRow2(0) = "Left" : viewRow2(1) = "Right" : viewRow2(2) = "Top"
        Dim imgCols(2) As Integer : imgCols(0) = 1 : imgCols(1) = 3 : imgCols(2) = 5

        Dim v As Integer
        For v = 0 To 2
            If gd.Exists(viewRow1(v)) Then
                PlaceImage wsh, gd(viewRow1(v)), _
                    startRow + 2, imgCols(v), BLOCK_IMG_H
            End If
        Next v
        For v = 0 To 2
            If gd.Exists(viewRow2(v)) Then
                PlaceImage wsh, gd(viewRow2(v)), _
                    startRow + 9, imgCols(v), BLOCK_IMG_H
            End If
        Next v
    Next blockIdx

    LogMessage "Report sheet filled with " & groupCount & " slide blocks."
    ThisWorkbook.Sheets("Report").Activate
End Sub

' ============================================================
' BuildTemplateBlock — format one slide block
' ============================================================
Private Sub BuildTemplateBlock(wsh As Worksheet, startRow As Long, _
    setID As Long, setLabel As String, analysisType As String, summaryValue As String)

    ' --- Title bar (row +0) ---
    With wsh.Range(wsh.Cells(startRow, 1), wsh.Cells(startRow, 6))
        .Merge
        .Value = setLabel & "  —  " & analysisType & "  |  " & summaryValue
        .Interior.Color = NAVY_COLOR
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .RowHeight = 22
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With

    ' Add tag in Col F for PPTXBuilder to find the block
    wsh.Cells(startRow, 6).Value = "BLOCK:" & setID

    ' --- View label rows ---
    Dim labels1(2) As String : labels1(0) = "ISO" : labels1(1) = "Front" : labels1(2) = "Back"
    Dim labels2(2) As String : labels2(0) = "Left" : labels2(1) = "Right" : labels2(2) = "Top"
    Dim imgCols(2) As Integer : imgCols(0) = 1 : imgCols(1) = 3 : imgCols(2) = 5

    Dim labelRow1 As Long : labelRow1 = startRow + 1
    Dim labelRow2 As Long : labelRow2 = startRow + 8

    Dim v As Integer
    For v = 0 To 2
        With wsh.Cells(labelRow1, imgCols(v))
            .Value = labels1(v)
            .Font.Bold = True
            .Font.Size = 9
            .Interior.Color = RGB(220, 230, 241)
            .HorizontalAlignment = xlCenter
        End With
        With wsh.Cells(labelRow2, imgCols(v))
            .Value = labels2(v)
            .Font.Bold = True
            .Font.Size = 9
            .Interior.Color = RGB(220, 230, 241)
            .HorizontalAlignment = xlCenter
        End With
    Next v

    ' Set row heights for image rows
    Dim r As Long
    For r = startRow + 2 To startRow + 7
        wsh.Rows(r).RowHeight = 48
    Next r
    For r = startRow + 9 To startRow + 14
        wsh.Rows(r).RowHeight = 48
    Next r

    ' --- Footer value (row +15) ---
    With wsh.Range(wsh.Cells(startRow + 15, 1), wsh.Cells(startRow + 15, 6))
        .Merge
        .Value = analysisType & ": " & summaryValue
        .Interior.Color = LIGHT_GRAY
        .Font.Bold = False
        .Font.Size = 10
        .RowHeight = 18
    End With

    ' --- Separator (row +16) ---
    With wsh.Rows(startRow + 16)
        .RowHeight = 8
        .Interior.Color = RGB(200, 200, 200)
    End With
End Sub

' ============================================================
' PlaceImage — insert image into sheet at target cell
' ============================================================
Private Sub PlaceImage(wsh As Worksheet, imgPath As String, _
    topRow As Long, leftCol As Long, rowSpan As Long)

    If Not FileExists(imgPath) Then Exit Sub

    Dim topCell    As Range : Set topCell    = wsh.Cells(topRow, leftCol)
    Dim bottomCell As Range : Set bottomCell = wsh.Cells(topRow + rowSpan - 1, leftCol)

    Dim L As Single : L = topCell.Left + 2
    Dim T As Single : T = topCell.Top + 2
    Dim W As Single : W = topCell.Width - 4
    Dim H As Single : H = bottomCell.Top + bottomCell.Height - topCell.Top - 4

    On Error Resume Next
    Dim pic As Object
    Set pic = wsh.Shapes.AddPicture( _
        Filename:=imgPath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=L, Top:=T, Width:=W, Height:=H)
    If Not pic Is Nothing Then
        pic.LockAspectRatio = msoFalse
        pic.Width = W
        pic.Height = H
    End If
    On Error GoTo 0
End Sub

' ============================================================
' ClearAllImages — remove all Shape objects from Report sheet
' ============================================================
Public Sub ClearAllImages()
    Dim wsh As Worksheet : Set wsh = ThisWorkbook.Sheets("Report")
    Dim shp As Shape
    Dim toDelete() As String
    Dim cnt As Integer : cnt = 0
    ReDim toDelete(0 To wsh.Shapes.Count)

    For Each shp In wsh.Shapes
        ' Don't delete buttons (Form Controls / ActiveX)
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            toDelete(cnt) = shp.Name
            cnt = cnt + 1
        End If
    Next shp

    Dim i As Integer
    For i = 0 To cnt - 1
        wsh.Shapes(toDelete(i)).Delete
    Next i

    LogMessage "Cleared " & cnt & " images from Report sheet."
End Sub

' Public overload (no args) for button binding
Public Sub ClearAllImagesBtn()
    ClearAllImages
End Sub

' ============================================================
' FillFromLastRun — shortcut: find manifest in output folder
' ============================================================
Public Sub FillFromLastRun()
    Dim cfg As PipelineCfg : cfg = ReadConfig()
    Dim manifestPath As String
    manifestPath = Replace(cfg.OutputFolder, "/", "\") & "\screenshot_manifest.txt"
    FillReportFromManifest manifestPath
End Sub
