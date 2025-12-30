Attribute VB_Name = "Module1"
Option Explicit

Public Sub Run_DataHygiene()

    Dim startTime As Double
    Dim elapsedTime As Double

    Dim wsCtrl As Worksheet
    Dim lastRowData As Long
    Dim lastColData As Long
    Dim rowsWritten As Long
    Dim colsWritten As Long

    Dim rngDataArea As Range
    Dim f As Range

    startTime = Timer

    Set wsCtrl = ThisWorkbook.Worksheets("Control")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Call Step1_TrimWhiteSpace
    Call Step2_StandardizeTextCase
    Call Step3_FillBlanksWithND
    Call Step4_FlagDuplicateRows

    Set rngDataArea = wsCtrl.Range("E:XFD")

    Set f = rngDataArea.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        lastRowData = 0
    Else
        lastRowData = f.Row
    End If

    Set f = rngDataArea.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If f Is Nothing Then
        lastColData = 0
    Else
        lastColData = f.Column
    End If

    If lastRowData >= 2 Then
        rowsWritten = lastRowData - 1
    Else
        rowsWritten = 0
    End If

    If lastColData >= 5 Then
        colsWritten = lastColData - 4
    Else
        colsWritten = 0
    End If

    With wsCtrl.Range("C13")
        .NumberFormat = "General"
        .Value = rowsWritten
    End With

    With wsCtrl.Range("C14")
        .NumberFormat = "General"
        .Value = colsWritten
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    elapsedTime = Timer - startTime

    With wsCtrl.Range("C15")
        .Value = Round(elapsedTime, 2)
        .NumberFormat = "0.00"" seconds"""
    End With

End Sub

Public Sub Step1_TrimWhiteSpace()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range
    Dim dstRange As Range
    Dim dataArr As Variant

    Dim r As Long, c As Long
    Dim originalText As String
    Dim workingText As String
    Dim beforeLen As Long, afterLen As Long
    Dim removedCount As Long

    Set wsSrc = ThisWorkbook.Worksheets("Data Input")
    Set wsDst = ThisWorkbook.Worksheets("Control")

    lastRow = wsSrc.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = wsSrc.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Set srcRange = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))
    dataArr = srcRange.Value

    removedCount = 0

    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)

            If VarType(dataArr(r, c)) = vbString Then
                originalText = dataArr(r, c)

                workingText = Replace(originalText, Chr(160), " ")
                beforeLen = Len(workingText)
                workingText = Trim(workingText)

                Do While InStr(workingText, "  ") > 0
                    workingText = Replace(workingText, "  ", " ")
                Loop

                afterLen = Len(workingText)
                removedCount = removedCount + (beforeLen - afterLen)

                dataArr(r, c) = workingText
            End If

        Next c
    Next r

    Set dstRange = wsDst.Range("E1").Resize(UBound(dataArr, 1), UBound(dataArr, 2))
    dstRange.Value = dataArr

    wsDst.Range("C19").Value = removedCount

End Sub

Public Sub Step2_StandardizeTextCase()

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range
    Dim dataArr As Variant

    Dim caseOption As String
    Dim r As Long, c As Long

    Set ws = ThisWorkbook.Worksheets("Control")
    caseOption = Trim(ws.Range("C8").Value)

    If caseOption = "None" Or caseOption = "" Then Exit Sub

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If lastRow < 2 Or lastCol < 5 Then Exit Sub

    Set srcRange = ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, lastCol))
    dataArr = srcRange.Value

    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)

            If VarType(dataArr(r, c)) = vbString Then
                Select Case caseOption
                    Case "Proper (Aa)"
                        dataArr(r, c) = StrConv(dataArr(r, c), vbProperCase)
                    Case "Lower (aa)"
                        dataArr(r, c) = LCase$(dataArr(r, c))
                    Case "Upper (AA)"
                        dataArr(r, c) = UCase$(dataArr(r, c))
                End Select
            End If

        Next c
    Next r

    srcRange.Value = dataArr

End Sub

Public Sub Step3_FillBlanksWithND()

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range
    Dim dataArr As Variant

    Dim rawOption As Variant
    Dim applyEnabled As Boolean

    Dim r As Long, c As Long
    Dim s As String
    Dim sNorm As String

    Set ws = ThisWorkbook.Worksheets("Control")

    rawOption = ws.Range("C9").Value
    applyEnabled = False

    If VarType(rawOption) = vbString Then
        s = Replace(CStr(rawOption), Chr(160), " ")
        If UCase$(Trim$(s)) = "YES" Then applyEnabled = True
    ElseIf IsNumeric(rawOption) Then
        If CLng(rawOption) = 1 Then applyEnabled = True
    End If

    If Not applyEnabled Then Exit Sub

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If lastRow < 2 Or lastCol < 5 Then Exit Sub

    Set srcRange = ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, lastCol))
    dataArr = srcRange.Value

    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)

            If IsEmpty(dataArr(r, c)) Then
                dataArr(r, c) = "N/D"
            ElseIf VarType(dataArr(r, c)) = vbString Then
                sNorm = Replace$(CStr(dataArr(r, c)), Chr$(160), " ")
                If Len(Trim$(sNorm)) = 0 Then
                    dataArr(r, c) = "N/D"
                End If
            End If

        Next c
    Next r

    srcRange.Value = dataArr

End Sub

Public Sub Step4_FlagDuplicateRows()

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim dataArr As Variant
    Dim dict As Object

    Dim keyColsInput As String
    Dim parts() As String
    Dim keyCols() As Long
    Dim hasKeyCols As Boolean

    Dim rawRemove As Variant
    Dim removeEnabled As Boolean
    Dim tmp As String

    Dim r As Long, c As Long, i As Long
    Dim token As String
    Dim colLetters As String
    Dim colNum As Long
    Dim ch As Long

    Dim key As String
    Dim v As Variant
    Dim s As String
    Dim partText As String

    Dim rowRange As Range
    Dim dataRange As Range

    Dim uniqueCount As Long
    Dim duplicateCount As Long
    Dim blankIdRows As Long
    Dim anyNonBlank As Boolean

    Dim totalCols As Long
    Dim maxRows As Long
    Dim uniqueArr() As Variant
    Dim dupArr() As Variant
    Dim uRow As Long, dRow As Long

    Dim dupSheet As Worksheet
    Dim sheetName As String

    Const DATA_FIRST_ROW As Long = 2
    Const DATA_FIRST_COL As Long = 5

    Set ws = ThisWorkbook.Worksheets("Control")

    rawRemove = ws.Range("C10").Value
    removeEnabled = False
    If VarType(rawRemove) = vbString Then
        tmp = Replace$(CStr(rawRemove), Chr$(160), " ")
        If UCase$(Trim$(tmp)) = "YES" Then removeEnabled = True
    ElseIf IsNumeric(rawRemove) Then
        If CLng(rawRemove) = 1 Then removeEnabled = True
    End If

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If lastRow < DATA_FIRST_ROW Or lastCol < DATA_FIRST_COL Then Exit Sub

    totalCols = lastCol - DATA_FIRST_COL + 1
    Set dataRange = ws.Range(ws.Cells(DATA_FIRST_ROW, DATA_FIRST_COL), ws.Cells(lastRow, lastCol))

    dataRange.Interior.Pattern = xlNone
    dataRange.Font.ColorIndex = xlColorIndexAutomatic
    dataRange.Borders.LineStyle = xlNone

    dataArr = ws.Range(ws.Cells(1, DATA_FIRST_COL), ws.Cells(lastRow, lastCol)).Value

    keyColsInput = CStr(ws.Range("B5").Value)
    keyColsInput = Replace$(keyColsInput, Chr$(160), " ")
    keyColsInput = UCase$(Trim$(keyColsInput))

    hasKeyCols = False

    If Len(keyColsInput) > 0 Then
        parts = Split(keyColsInput, ",")
        ReDim keyCols(0 To UBound(parts))

        For i = 0 To UBound(parts)
            token = Trim$(parts(i))
            If Len(token) > 0 Then
                colLetters = token
                colNum = 0

                For c = 1 To Len(colLetters)
                    ch = Asc(Mid$(colLetters, c, 1))
                    If ch < 65 Or ch > 90 Then
                        colNum = 0
                        Exit For
                    End If
                    colNum = (colNum * 26) + (ch - 64)
                Next c

                If colNum > 0 Then
                    If colNum < DATA_FIRST_COL Then
                        colNum = DATA_FIRST_COL + (colNum - 1)
                    End If
                    keyCols(i) = colNum
                    hasKeyCols = True
                End If
            End If
        Next i
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    uniqueCount = 0
    duplicateCount = 0
    blankIdRows = 0

    If removeEnabled Then
        maxRows = UBound(dataArr, 1)
        ReDim uniqueArr(1 To maxRows, 1 To totalCols)
        ReDim dupArr(1 To maxRows, 1 To totalCols)

        For c = 1 To totalCols
            uniqueArr(1, c) = dataArr(1, c)
            dupArr(1, c) = dataArr(1, c)
        Next c

        uRow = 1
        dRow = 1
    End If

    For r = DATA_FIRST_ROW To UBound(dataArr, 1)

        key = vbNullString
        anyNonBlank = False

        If hasKeyCols Then
            For i = LBound(keyCols) To UBound(keyCols)
                If keyCols(i) >= DATA_FIRST_COL And keyCols(i) <= lastCol Then
                    v = dataArr(r, (keyCols(i) - DATA_FIRST_COL) + 1)

                    If IsError(v) Then
                        partText = "#ERR"
                        anyNonBlank = True
                    ElseIf IsEmpty(v) Then
                        partText = vbNullString
                    ElseIf VarType(v) = vbString Then
                        s = Replace$(CStr(v), Chr$(160), " ")
                        s = Trim$(s)
                        partText = UCase$(s)
                        If Len(s) > 0 Then anyNonBlank = True
                    Else
                        partText = CStr(v)
                        If Len(partText) > 0 Then anyNonBlank = True
                    End If

                    key = key & partText & ChrW$(30)
                End If
            Next i

            If Not anyNonBlank Then blankIdRows = blankIdRows + 1
        Else
            For c = 1 To totalCols
                v = dataArr(r, c)

                If IsError(v) Then
                    partText = "#ERR"
                ElseIf IsEmpty(v) Then
                    partText = vbNullString
                ElseIf VarType(v) = vbString Then
                    s = Replace$(CStr(v), Chr$(160), " ")
                    partText = UCase$(Trim$(s))
                Else
                    partText = CStr(v)
                End If

                key = key & partText & ChrW$(30)
            Next c
        End If

        If dict.Exists(key) Then
            duplicateCount = duplicateCount + 1

            If removeEnabled Then
                dRow = dRow + 1
                For c = 1 To totalCols
                    dupArr(dRow, c) = dataArr(r, c)
                Next c
            Else
                Set rowRange = ws.Range(ws.Cells(r, DATA_FIRST_COL), ws.Cells(r, lastCol))
                rowRange.Style = "Bad"
                rowRange.Borders.LineStyle = xlContinuous
                rowRange.Borders.Weight = xlThin
                rowRange.Borders.Color = RGB(200, 200, 200)
            End If
        Else
            dict.Add key, True
            uniqueCount = uniqueCount + 1

            If removeEnabled Then
                uRow = uRow + 1
                For c = 1 To totalCols
                    uniqueArr(uRow, c) = dataArr(r, c)
                Next c
            End If
        End If

    Next r

    If removeEnabled Then
        ws.Range(ws.Cells(1, DATA_FIRST_COL), ws.Cells(lastRow, lastCol)).ClearContents
        ws.Range(ws.Cells(1, DATA_FIRST_COL), ws.Cells(uRow, lastCol)).Value = uniqueArr

        sheetName = "Duplicates"
        On Error Resume Next
        Set dupSheet = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo 0

        If dupSheet Is Nothing Then
            Set dupSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            dupSheet.Name = sheetName
        Else
            dupSheet.Cells.Clear
        End If

        If dRow >= 1 Then
            dupSheet.Range(dupSheet.Cells(1, 1), dupSheet.Cells(dRow, totalCols)).Value = dupArr
        End If
    End If

    ws.Range("C16").Value = duplicateCount
    ws.Range("C17").Value = uniqueCount

    If hasKeyCols Then
        ws.Range("C18").Value = blankIdRows
    Else
        ws.Range("C18").Value = "N/A"
    End If

End Sub

