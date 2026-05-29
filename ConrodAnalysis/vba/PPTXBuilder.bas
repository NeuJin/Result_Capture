Attribute VB_Name = "PPTXBuilder"
Option Explicit

' ============================================================
' PPTXBuilder — Read Report sheet layout → generate PPTX
' via PowerPoint COM automation.
'
' Slide layout (widescreen 13.33" × 7.5" = 960×540 pt):
'   - Title bar  : 0–22pt      dark blue, white text
'   - Image grid : 22–504pt    6 images in 2 rows of 3
'   - Footer     : 504–540pt   light gray, summary value
' ============================================================

Private Const SLIDE_W_PT    As Single = 960    ' 13.33" × 72
Private Const SLIDE_H_PT    As Single = 540    ' 7.5" × 72
Private Const TITLE_H_PT    As Single = 22
Private Const FOOTER_H_PT   As Single = 22
Private Const IMG_AREA_H_PT As Single = 496    ' 540 - 22 - 22
Private Const IMG_ROW_H_PT  As Single = 248    ' 496 / 2
Private Const IMG_COL_W_PT  As Single = 320    ' 960 / 3

Private Const NAVY_PPT      As Long = 3368601  ' RGB(51,102,153)
Private Const LIGHT_GRAY_PPT As Long = 15921906

Public Sub BuildFromReportSheet()
    Dim wsh As Worksheet : Set wsh = ThisWorkbook.Sheets("Report")
    Dim cfg As PipelineCfg : cfg = ReadConfig()

    ' Find all BLOCK: tags in Col F
    Dim blockRows() As Long
    Dim blockIDs()  As Long
    Dim blockCount  As Long : blockCount = 0
    ReDim blockRows(0 To 200)
    ReDim blockIDs(0 To 200)

    Dim r As Long
    Dim lastRow As Long : lastRow = wsh.Cells(wsh.Rows.Count, 1).End(xlUp).Row
    For r = DATA_START_ROW To lastRow
        Dim cellVal As String : cellVal = CStr(wsh.Cells(r, 6).Value)
        If Left(cellVal, 6) = "BLOCK:" Then
            blockRows(blockCount) = r
            blockIDs(blockCount)  = CLng(Mid(cellVal, 7))
            blockCount = blockCount + 1
        End If
    Next r

    If blockCount = 0 Then
        MsgBox "No slide blocks found in Report sheet." & vbLf & _
               "Run analysis first or click [Auto-fill from last run].", vbExclamation
        Exit Sub
    End If

    ' --- Open PowerPoint ---
    Dim ppt As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then
        Set ppt = CreateObject("PowerPoint.Application")
    End If
    ppt.Visible = False

    Dim prs As Object : Set prs = ppt.Presentations.Add(False)

    ' Slide dimensions
    prs.PageSetup.SlideWidth  = SLIDE_W_PT
    prs.PageSetup.SlideHeight = SLIDE_H_PT

    ' --- Build one slide per block ---
    Dim bi As Long
    For bi = 0 To blockCount - 1
        Dim startRow As Long : startRow = blockRows(bi)

        ' Read title text (Col A of title row — stripped of the BLOCK tag)
        Dim titleText As String
        titleText = CStr(wsh.Cells(startRow, 1).Value)

        ' Read footer text
        Dim footerText As String
        footerText = CStr(wsh.Cells(startRow + 15, 1).Value)

        ' Add blank slide
        Dim sld As Object
        Set sld = prs.Slides.Add(prs.Slides.Count + 1, 12)  ' 12 = ppLayoutBlank

        ' Remove any default placeholders
        Dim shp As Object
        For Each shp In sld.Shapes
            shp.Delete
        Next shp

        ' --- Title bar ---
        AddRectangle sld, 0, 0, SLIDE_W_PT, TITLE_H_PT, NAVY_PPT
        AddLabel sld, titleText, 4, 2, SLIDE_W_PT - 8, TITLE_H_PT, _
            RGB(255, 255, 255), 12, True, xlLeft

        ' --- Footer ---
        AddRectangle sld, 0, SLIDE_H_PT - FOOTER_H_PT, SLIDE_W_PT, FOOTER_H_PT, LIGHT_GRAY_PPT
        AddLabel sld, footerText, 4, SLIDE_H_PT - FOOTER_H_PT + 3, _
            SLIDE_W_PT - 8, FOOTER_H_PT, RGB(50, 50, 50), 9, False, xlCenter

        ' --- Images ---
        Dim viewRow1(2) As String : viewRow1(0) = "ISO"  : viewRow1(1) = "Front" : viewRow1(2) = "Back"
        Dim viewRow2(2) As String : viewRow2(0) = "Left" : viewRow2(1) = "Right" : viewRow2(2) = "Top"
        Dim imgCols(2)  As Integer : imgCols(0) = 1 : imgCols(1) = 3 : imgCols(2) = 5

        Dim v As Integer
        For v = 0 To 2
            ' Row 1 images (starting at row startRow+2)
            Dim imgPath1 As String
            imgPath1 = GetImagePathFromShape(wsh, startRow + 2, imgCols(v))
            If imgPath1 <> "" Then
                AddPicture sld, imgPath1, _
                    v * IMG_COL_W_PT, TITLE_H_PT, IMG_COL_W_PT, IMG_ROW_H_PT
            End If

            ' Row 2 images (starting at row startRow+9)
            Dim imgPath2 As String
            imgPath2 = GetImagePathFromShape(wsh, startRow + 9, imgCols(v))
            If imgPath2 <> "" Then
                AddPicture sld, imgPath2, _
                    v * IMG_COL_W_PT, TITLE_H_PT + IMG_ROW_H_PT, IMG_COL_W_PT, IMG_ROW_H_PT
            End If
        Next v

        ' --- View name labels (small, on top of images) ---
        Dim viewNames(5) As String
        viewNames(0) = "ISO" : viewNames(1) = "Front" : viewNames(2) = "Back"
        viewNames(3) = "Left" : viewNames(4) = "Right" : viewNames(5) = "Top"
        Dim vi As Integer
        For vi = 0 To 5
            Dim vRow As Integer : vRow = vi \ 3
            Dim vCol As Integer : vCol = vi Mod 3
            AddLabel sld, viewNames(vi), _
                vCol * IMG_COL_W_PT + 4, _
                TITLE_H_PT + vRow * IMG_ROW_H_PT + 2, _
                60, 14, RGB(255, 255, 255), 8, True, xlLeft
        Next vi
    Next bi

    ' --- Save ---
    Dim outputPath As String
    outputPath = Replace(cfg.OutputFolder, "/", "\") & "\reports\ConrodReport.pptx"

    ' Ensure reports folder
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Replace(cfg.OutputFolder, "/", "\") & "\reports") Then
        fso.CreateFolder Replace(cfg.OutputFolder, "/", "\") & "\reports"
    End If

    prs.SaveAs outputPath, 24  ' 24 = ppSaveAsOpenXMLPresentation
    prs.Close
    ppt.Quit

    LogMessage "PPTX saved: " & outputPath
    MsgBox "PPTX generated successfully!" & vbLf & outputPath, vbInformation, "Report Done"

    ' Open the file
    Shell "explorer.exe " & Chr(34) & outputPath & Chr(34), vbNormalFocus
End Sub

' ============================================================
' Helper: find the source path of a picture shape near a cell
' ============================================================
Private Function GetImagePathFromShape(wsh As Worksheet, _
    topRow As Long, leftCol As Long) As String

    Dim targetCell As Range : Set targetCell = wsh.Cells(topRow, leftCol)
    Dim shp As Shape

    For Each shp In wsh.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' Check if shape overlaps with target cell
            If shp.TopLeftCell.Row >= topRow And _
               shp.TopLeftCell.Row <= topRow + 5 And _
               shp.TopLeftCell.Column = leftCol Then
                ' Return the shape's AlternativeText as path (set when inserting)
                GetImagePathFromShape = shp.AlternativeText
                Exit Function
            End If
        End If
    Next shp
    GetImagePathFromShape = ""
End Function

' ============================================================
' PPT shape helpers (use Points, not inches)
' ============================================================
Private Sub AddRectangle(sld As Object, L As Single, T As Single, _
    W As Single, H As Single, fillColor As Long)

    Dim shp As Object
    Set shp = sld.Shapes.AddShape(1, L, T, W, H)  ' 1 = msoShapeRectangle
    shp.Fill.ForeColor.RGB = fillColor
    shp.Fill.Solid
    shp.Line.Visible = False
End Sub

Private Sub AddLabel(sld As Object, txt As String, _
    L As Single, T As Single, W As Single, H As Single, _
    fontColor As Long, fontSize As Integer, bold As Boolean, align As Integer)

    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, L, T, W, H)
    With shp.TextFrame
        .WordWrap = True
        .AutoSize = 0  ' ppAutoSizeNone
    End With
    With shp.TextFrame.TextRange
        .Text = txt
        .Font.Color.RGB = fontColor
        .Font.Size = fontSize
        .Font.Bold = bold
        .ParagraphFormat.Alignment = align
    End With
    shp.Line.Visible = False
    shp.Fill.Visible = False
End Sub

Private Sub AddPicture(sld As Object, imgPath As String, _
    L As Single, T As Single, W As Single, H As Single)

    On Error Resume Next
    sld.Shapes.AddPicture imgPath, False, True, L, T, W, H
    On Error GoTo 0
End Sub
