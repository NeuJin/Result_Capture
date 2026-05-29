Attribute VB_Name = "InpParser"
Option Explicit

' ============================================================
' InpParser — Reads Abaqus .inp file, extracts *NSET sections
' Populates the NodeSets sheet with NSET names + node counts
' ============================================================

Public Sub ParseInpFile()
    Dim filePath As String
    filePath = ThisWorkbook.Sheets("Config").Range("B4").Value

    If filePath = "" Then
        MsgBox "Please enter the .inp model file path in Config sheet (B4).", vbExclamation
        Exit Sub
    End If

    If Not FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    Dim nsets As Object  ' Scripting.Dictionary: name -> comma-separated node IDs
    Set nsets = ParseNsets(filePath)

    If nsets.Count = 0 Then
        MsgBox "No *NSET found in file: " & filePath, vbExclamation
        Exit Sub
    End If

    ' Write to NodeSets sheet
    Dim wsh As Worksheet
    Set wsh = ThisWorkbook.Sheets("NodeSets")

    ' Clear existing data (keep header row 1)
    Dim lastRow As Long
    lastRow = wsh.Cells(wsh.Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        wsh.Rows("2:" & lastRow).ClearContents
    End If

    Dim r As Long : r = 2
    Dim key As Variant
    Dim setID As Long : setID = 1

    For Each key In nsets.Keys
        Dim nodeList As String : nodeList = nsets(key)
        Dim nodeCount As Long : nodeCount = CountCSV(nodeList)

        wsh.Cells(r, 1).Value = setID           ' Col A: Set ID
        wsh.Cells(r, 2).Value = CStr(key)       ' Col B: NSET Name
        wsh.Cells(r, 3).Value = "Y"             ' Col C: Include (default Y)
        wsh.Cells(r, 4).Value = CStr(key)       ' Col D: Display Label (default = name)
        wsh.Cells(r, 5).Value = nodeCount       ' Col E: Node count (info)
        wsh.Cells(r, 6).Value = nodeList        ' Col F: Node IDs (hidden helper col)

        r = r + 1
        setID = setID + 1
    Next key

    ' Auto-fit columns
    wsh.Columns("A:E").AutoFit

    MsgBox "Parsed " & nsets.Count & " NSETs from .inp file.", vbInformation
End Sub

' ============================================================
' Core parser — returns Dictionary {nset_name -> "id1,id2,..."}
' Handles: regular list, GENERATE (start,end,step), multi-line,
'          ** comments, mixed-case *NSET keyword
' ============================================================
Public Function ParseNsets(filePath As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim f As Object
    Set f = fso.OpenTextFile(filePath, 1)  ' ForReading

    Dim currentName As String : currentName = ""
    Dim isGenerate As Boolean : isGenerate = False
    Dim nodeBuffer As String : nodeBuffer = ""

    Do While Not f.AtEndOfStream
        Dim line As String
        line = Trim(f.ReadLine)

        ' Skip blank lines and comments
        If line = "" Or Left(line, 2) = "**" Then GoTo NextLine

        ' Detect *NSET keyword (case-insensitive)
        If UCase(Left(Replace(line, " ", ""), 5)) = "*NSET" Then
            ' Save previous NSET if any
            If currentName <> "" Then
                If isGenerate Then
                    Call SaveGeneratedNset(result, currentName, nodeBuffer)
                Else
                    result(currentName) = Trim(nodeBuffer, ",")
                End If
            End If

            ' Parse new NSET name
            currentName = ExtractNsetName(line)
            isGenerate = (InStr(UCase(line), "GENERATE") > 0)
            nodeBuffer = ""
            GoTo NextLine
        End If

        ' Detect any other keyword (*ELEMENT, *NODE, etc.) — end current NSET
        If Left(line, 1) = "*" And Left(line, 2) <> "**" Then
            If currentName <> "" Then
                If isGenerate Then
                    Call SaveGeneratedNset(result, currentName, nodeBuffer)
                Else
                    result(currentName) = TrimComma(nodeBuffer)
                End If
                currentName = ""
                nodeBuffer = ""
            End If
            GoTo NextLine
        End If

        ' Accumulate node data (only if inside a NSET block)
        If currentName <> "" Then
            nodeBuffer = nodeBuffer & line
            ' Multi-line: if line ends with comma, more data follows → continue
            ' If not → still continue reading (could be last line of set)
        End If

NextLine:
    Loop

    f.Close

    ' Save last NSET
    If currentName <> "" Then
        If isGenerate Then
            Call SaveGeneratedNset(result, currentName, nodeBuffer)
        Else
            result(currentName) = TrimComma(nodeBuffer)
        End If
    End If

    Set ParseNsets = result
End Function

' Extract NSET name from *NSET,NSET=BigEnd or *NSET, NSET=BigEnd
Private Function ExtractNsetName(line As String) As String
    Dim parts() As String
    parts = Split(line, ",")
    Dim i As Integer
    For i = 0 To UBound(parts)
        Dim part As String : part = Trim(parts(i))
        If UCase(Left(part, 5)) = "NSET=" Then
            ExtractNsetName = Trim(Mid(part, 6))
            Exit Function
        End If
    Next i
    ExtractNsetName = "UNKNOWN_" & CStr(Int(Rnd() * 9999))
End Function

' Expand GENERATE: "start, end, step" → comma-separated IDs
Private Sub SaveGeneratedNset(result As Object, nsetName As String, raw As String)
    Dim parts() As String
    parts = Split(Replace(raw, " ", ""), ",")

    If UBound(parts) < 1 Then Exit Sub  ' malformed

    Dim startID As Long : startID = CLng(parts(0))
    Dim endID As Long   : endID   = CLng(parts(1))
    Dim stepVal As Long : stepVal = 1
    If UBound(parts) >= 2 And Trim(parts(2)) <> "" Then
        stepVal = CLng(parts(2))
    End If

    Dim ids As String : ids = ""
    Dim id As Long
    For id = startID To endID Step stepVal
        If ids = "" Then
            ids = CStr(id)
        Else
            ids = ids & "," & CStr(id)
        End If
    Next id

    result(nsetName) = ids
End Sub

' Count items in comma-separated string
Private Function CountCSV(s As String) As Long
    If Trim(s) = "" Then
        CountCSV = 0
        Exit Function
    End If
    CountCSV = UBound(Split(s, ",")) + 1
End Function

' Remove leading/trailing commas and spaces
Private Function TrimComma(s As String) As String
    s = Replace(s, " ", "")
    Do While Left(s, 1) = ","
        s = Mid(s, 2)
    Loop
    Do While Right(s, 1) = ","
        s = Left(s, Len(s) - 1)
    Loop
    TrimComma = s
End Function

' Check file exists
Public Function FileExists(path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function
