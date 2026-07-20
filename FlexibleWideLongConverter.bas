Attribute VB_Name = "FlexibleWideLongConverter"
Option Explicit

' Converts between wide (multiple measurement columns) and long format.
'
' Wide -> Long input:
' ID column(s) | Variable 1 | Variable 2 | ...
'
' Long -> Wide input:
' ID column(s) | Variable | Value
'
' The user specifies the conversion direction and number of ID columns.

Public Sub ConvertWideAndLongFormats()
    Dim src As Range, directionChoice As Variant, idChoice As Variant
    Dim nID As Long, nCols As Long

    On Error Resume Next
    Set src = Application.InputBox( _
        Prompt:="Select the complete table INCLUDING headers.", _
        Title:="Wide / Long Format Converter", Type:=8)
    On Error GoTo CleanFail

    If src Is Nothing Then Exit Sub
    If src.Areas.Count > 1 Then Err.Raise vbObjectError + 1, , "Select one contiguous range."
    If src.Rows.Count < 2 Then Err.Raise vbObjectError + 2, , "The selection needs a header and at least one data row."

    directionChoice = Application.InputBox( _
        Prompt:="Choose the conversion:" & vbCrLf & vbCrLf & _
               "1 = Wide to Long" & vbCrLf & _
               "    ID column(s) | Outcome 1 | Outcome 2 | ..." & vbCrLf & vbCrLf & _
               "2 = Long to Wide" & vbCrLf & _
               "    ID column(s) | Variable | Value", _
        Title:="Choose conversion direction", Default:="1", Type:=2)
    If directionChoice = False Then Exit Sub

    nCols = src.Columns.Count
    idChoice = Application.InputBox( _
        Prompt:="How many identifier/grouping columns are at the LEFT of the table?" & vbCrLf & vbCrLf & _
               "Examples:" & vbCrLf & _
               "Subject | Group | Time1 | Time2 = 2 ID columns" & vbCrLf & _
               "Subject | Group | Time | Value = 2 ID columns", _
        Title:="Number of identifier columns", Default:="1", Type:=1)
    If idChoice = False Then Exit Sub
    nID = CLng(idChoice)
    If nID < 1 Then Err.Raise vbObjectError + 3, , "The number of identifier columns must be at least 1."

    Select Case Trim$(CStr(directionChoice))
        Case "1"
            If nID >= nCols Then Err.Raise vbObjectError + 4, , "Wide-to-long conversion needs at least one measurement column after the ID columns."
            ConvertWideToLong src, nID
        Case "2"
            If nCols <> nID + 2 Then Err.Raise vbObjectError + 5, , _
                "For long-to-wide conversion, select ID column(s), then exactly one Variable column and one Value column."
            ConvertLongToWide src, nID
        Case Else
            Err.Raise vbObjectError + 6, , "Enter 1 or 2."
    End Select
    Exit Sub

CleanFail:
    MsgBox "Could not convert the data: " & Err.Description, vbExclamation
End Sub

Private Sub ConvertWideToLong(ByVal src As Range, ByVal nID As Long)
    Dim data As Variant, wb As Workbook, wsOut As Worksheet
    Dim r As Long, c As Long, outRow As Long, lastCol As Long
    Dim variableHeader As Variant, valueHeader As Variant
    Dim includeBlank As VbMsgBoxResult

    data = src.Value2
    lastCol = UBound(data, 2)
    Set wb = src.Worksheet.Parent

    variableHeader = Application.InputBox( _
        Prompt:="Name the new column containing the former measurement headers.", _
        Title:="Variable-column name", Default:="Variable", Type:=2)
    If variableHeader = False Then Exit Sub

    valueHeader = Application.InputBox( _
        Prompt:="Name the new column containing the measurement values.", _
        Title:="Value-column name", Default:="Value", Type:=2)
    If valueHeader = False Then Exit Sub

    includeBlank = MsgBox( _
        "Include blank measurement cells as rows in the long table?" & vbCrLf & vbCrLf & _
        "Yes = retain explicit missing observations" & vbCrLf & _
        "No = omit blank observations", _
        vbYesNoCancel + vbQuestion, "Blank measurements")
    If includeBlank = vbCancel Then Exit Sub

    Application.ScreenUpdating = False
    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = UniqueConverterSheetName(wb, "Converted Long")

    For c = 1 To nID
        wsOut.Cells(1, c).Value = data(1, c)
    Next c
    wsOut.Cells(1, nID + 1).Value = CStr(variableHeader)
    wsOut.Cells(1, nID + 2).Value = CStr(valueHeader)

    outRow = 2
    For r = 2 To UBound(data, 1)
        For c = nID + 1 To lastCol
            If includeBlank = vbYes Or Not IsBlankLike(data(r, c)) Then
                wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, nID)).Value = _
                    src.Worksheet.Range(src.Cells(r, 1), src.Cells(r, nID)).Value
                wsOut.Cells(outRow, nID + 1).Value = data(1, c)
                If Not IsBlankLike(data(r, c)) Then wsOut.Cells(outRow, nID + 2).Value = data(r, c)
                outRow = outRow + 1
            End If
        Next c
    Next r

    FormatConvertedSheet wsOut, outRow - 1, nID + 2
    wsOut.Range("A" & outRow + 1).Value = "Source: " & src.Worksheet.Name & "!" & src.Address(False, False)
    wsOut.Activate
    Application.ScreenUpdating = True
    MsgBox "Wide-to-long conversion completed on '" & wsOut.Name & "'.", vbInformation
End Sub

Private Sub ConvertLongToWide(ByVal src As Range, ByVal nID As Long)
    Dim data As Variant, wb As Workbook, wsOut As Worksheet
    Dim variableCol As Long, valueCol As Long
    Dim idSeen As Object, variableSeen As Object, cellValues As Object, counts As Object
    Dim idOrder As Collection, variableOrder As Collection
    Dim r As Long, c As Long, i As Long, j As Long, outRow As Long
    Dim idKey As String, variableName As String, cellKey As String
    Dim displayIDs As Object, idArray As Variant, v As Variant
    Dim duplicateMode As Variant, duplicateCount As Long

    data = src.Value2
    variableCol = nID + 1
    valueCol = nID + 2
    Set wb = src.Worksheet.Parent

    duplicateMode = Application.InputBox( _
        Prompt:="If the same ID combination and Variable occurs more than once, choose:" & vbCrLf & vbCrLf & _
               "1 = Stop and report duplicates (safest default)" & vbCrLf & _
               "2 = Use the first value" & vbCrLf & _
               "3 = Average numeric duplicates", _
        Title:="Duplicate handling", Default:="1", Type:=2)
    If duplicateMode = False Then Exit Sub
    If InStr("123", Trim$(CStr(duplicateMode))) = 0 Then Err.Raise vbObjectError + 20, , "Enter 1, 2, or 3."

    Set idSeen = CreateObject("Scripting.Dictionary")
    Set variableSeen = CreateObject("Scripting.Dictionary")
    Set cellValues = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
    Set displayIDs = CreateObject("Scripting.Dictionary")
    Set idOrder = New Collection
    Set variableOrder = New Collection

    For r = 2 To UBound(data, 1)
        idKey = BuildIDKey(data, r, nID)
        variableName = Trim$(CStr(data(r, variableCol)))

        If Len(Replace(idKey, ChrW(30), "")) > 0 And Len(variableName) > 0 Then
            If Not idSeen.Exists(idKey) Then
                idSeen.Add idKey, True
                idOrder.Add idKey
                displayIDs.Add idKey, GetIDArray(data, r, nID)
            End If
            If Not variableSeen.Exists(variableName) Then
                variableSeen.Add variableName, True
                variableOrder.Add variableName
            End If

            If Not IsBlankLike(data(r, valueCol)) Then
                cellKey = idKey & ChrW(29) & variableName
                If Not cellValues.Exists(cellKey) Then
                    cellValues.Add cellKey, data(r, valueCol)
                    counts.Add cellKey, 1
                Else
                    duplicateCount = duplicateCount + 1
                    Select Case Trim$(CStr(duplicateMode))
                        Case "1"
                            Err.Raise vbObjectError + 21, , _
                                "Duplicate ID/Variable combinations were found. Example variable: '" & variableName & _
                                "'. Choose averaging or first-value handling, or resolve duplicates in the source data."
                        Case "2"
                            ' Keep the original value.
                        Case "3"
                            If Not IsNumeric(cellValues(cellKey)) Or Not IsNumeric(data(r, valueCol)) Then
                                Err.Raise vbObjectError + 22, , _
                                    "A duplicate contains nonnumeric values and cannot be averaged."
                            End If
                            cellValues(cellKey) = CDbl(cellValues(cellKey)) + CDbl(data(r, valueCol))
                            counts(cellKey) = CLng(counts(cellKey)) + 1
                    End Select
                End If
            End If
        End If
    Next r

    If idOrder.Count = 0 Or variableOrder.Count = 0 Then Err.Raise vbObjectError + 23, , "No valid ID/Variable rows were found."

    Application.ScreenUpdating = False
    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = UniqueConverterSheetName(wb, "Converted Wide")

    For c = 1 To nID
        wsOut.Cells(1, c).Value = data(1, c)
    Next c
    For j = 1 To variableOrder.Count
        wsOut.Cells(1, nID + j).Value = CStr(variableOrder(j))
    Next j

    outRow = 2
    For i = 1 To idOrder.Count
        idKey = CStr(idOrder(i))
        idArray = displayIDs(idKey)
        For c = 1 To nID
            wsOut.Cells(outRow, c).Value = idArray(c)
        Next c

        For j = 1 To variableOrder.Count
            variableName = CStr(variableOrder(j))
            cellKey = idKey & ChrW(29) & variableName
            If cellValues.Exists(cellKey) Then
                v = cellValues(cellKey)
                If Trim$(CStr(duplicateMode)) = "3" Then v = CDbl(v) / CLng(counts(cellKey))
                wsOut.Cells(outRow, nID + j).Value = v
            End If
        Next j
        outRow = outRow + 1
    Next i

    FormatConvertedSheet wsOut, outRow - 1, nID + variableOrder.Count
    wsOut.Range("A" & outRow + 1).Value = "Source: " & src.Worksheet.Name & "!" & src.Address(False, False)
    If duplicateCount > 0 Then
        wsOut.Range("A" & outRow + 2).Value = duplicateCount & " duplicate row(s) were handled using option " & CStr(duplicateMode) & "."
        wsOut.Range("A" & outRow + 2).Interior.Color = RGB(255, 235, 156)
    End If
    wsOut.Activate
    Application.ScreenUpdating = True
    MsgBox "Long-to-wide conversion completed on '" & wsOut.Name & "'.", vbInformation
End Sub

Private Function BuildIDKey(ByVal data As Variant, ByVal rowNumber As Long, ByVal nID As Long) As String
    Dim c As Long, result As String
    For c = 1 To nID
        If c > 1 Then result = result & ChrW(30)
        result = result & CStr(data(rowNumber, c))
    Next c
    BuildIDKey = result
End Function

Private Function GetIDArray(ByVal data As Variant, ByVal rowNumber As Long, ByVal nID As Long) As Variant
    Dim result() As Variant, c As Long
    ReDim result(1 To nID)
    For c = 1 To nID
        result(c) = data(rowNumber, c)
    Next c
    GetIDArray = result
End Function

Private Function IsBlankLike(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankLike = False
    Else
        IsBlankLike = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Sub FormatConvertedSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    ws.Rows(1).AutoFilter
    ws.Application.ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ws.Application.ActiveWindow.FreezePanes = True
    ws.Columns.AutoFit
End Sub

Private Function UniqueConverterSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, n As Long
    candidate = baseName
    Do While ConverterSheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(baseName, 25) & " " & n
    Loop
    UniqueConverterSheetName = candidate
End Function

Private Function ConverterSheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    ConverterSheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
