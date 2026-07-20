Attribute VB_Name = "RepeatedMeasuresLongFormatMeanSEM"
Option Explicit

' Creates a repeated-measures mean +/- SEM line graph from LONG-format data.
'
' Grouped input:
' Subject ID | Group | Time | Value
'
' Ungrouped input:
' Subject ID | Time | Value
'
' Each row is one observation. The macro graphs group means only;
' individual subject trajectories and raw values are not plotted.

Public Sub CreateRepeatedMeasuresLongFormatChart()
    Dim src As Range, wb As Workbook, wsOut As Worksheet
    Dim data As Variant, groupedAnswer As VbMsgBoxResult
    Dim isGrouped As Boolean, nRows As Long, nCols As Long
    Dim subjectCol As Long, groupCol As Long, timeCol As Long, valueCol As Long
    Dim groupSeen As Object, timeSeen As Object, subjectTimeSeen As Object
    Dim groupOrder As Collection, timeOrder As Collection, vals As Collection
    Dim r As Long, i As Long, j As Long
    Dim subjectID As String, groupName As String, timeName As String, obsKey As String
    Dim valueItem As Variant
    Dim outMeanStart As Long, outSemStart As Long, outNStart As Long
    Dim meanVal As Double, sdVal As Double, semVal As Double
    Dim validCount As Long, duplicateCount As Long
    Dim chartObj As ChartObject, ch As Chart, s As Series
    Dim timeRange As Range, meanRange As Range, semRange As Range
    Dim colors As Variant, clr As Long
    Dim yMin As Double, yMax As Double, yPad As Double, candidate As Double

    On Error Resume Next
    Set src = Application.InputBox( _
        Prompt:="Select the complete LONG-format repeated-measures table INCLUDING headers." & vbCrLf & vbCrLf & _
               "Grouped: Subject ID | Group | Time | Value" & vbCrLf & _
               "Ungrouped: Subject ID | Time | Value", _
        Title:="Long-format repeated-measures mean +/- SEM graph", Type:=8)
    On Error GoTo CleanFail

    If src Is Nothing Then Exit Sub
    If src.Areas.Count > 1 Then Err.Raise vbObjectError + 1, , "Select one contiguous range."

    nRows = src.Rows.Count
    nCols = src.Columns.Count
    If nRows < 3 Then Err.Raise vbObjectError + 2, , "The selection needs a header and at least two observation rows."
    If nCols <> 3 And nCols <> 4 Then Err.Raise vbObjectError + 3, , _
        "Select exactly 3 columns for ungrouped data or 4 columns for grouped data."

    If nCols = 4 Then
        groupedAnswer = MsgBox( _
            "Use column 2 as the grouping variable?" & vbCrLf & vbCrLf & _
            "Expected layout: Subject ID | Group | Time | Value", _
            vbYesNoCancel + vbQuestion, "Grouped long-format data?")
        If groupedAnswer = vbCancel Then Exit Sub
        If groupedAnswer = vbNo Then Err.Raise vbObjectError + 4, , _
            "For ungrouped data, select only Subject ID, Time, and Value."
        isGrouped = True
    Else
        isGrouped = False
    End If

    subjectCol = 1
    If isGrouped Then
        groupCol = 2: timeCol = 3: valueCol = 4
    Else
        groupCol = 0: timeCol = 2: valueCol = 3
    End If

    Set wb = src.Worksheet.Parent
    data = src.Value2
    Set groupSeen = CreateObject("Scripting.Dictionary")
    Set timeSeen = CreateObject("Scripting.Dictionary")
    Set subjectTimeSeen = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection
    Set timeOrder = New Collection

    ' Preserve first-appearance order for groups and time points.
    For r = 2 To nRows
        subjectID = Trim$(CStr(data(r, subjectCol)))
        timeName = Trim$(CStr(data(r, timeCol)))
        If isGrouped Then groupName = Trim$(CStr(data(r, groupCol))) Else groupName = "Mean"
        valueItem = data(r, valueCol)

        If Len(subjectID) > 0 And Len(groupName) > 0 And Len(timeName) > 0 And _
           IsNumeric(valueItem) And Len(CStr(valueItem)) > 0 Then

            If Not groupSeen.Exists(groupName) Then
                groupSeen.Add groupName, True
                groupOrder.Add groupName
            End If
            If Not timeSeen.Exists(timeName) Then
                timeSeen.Add timeName, True
                timeOrder.Add timeName
            End If

            ' Flag multiple numeric observations for the same subject/time combination.
            obsKey = subjectID & ChrW(30) & groupName & ChrW(30) & timeName
            If subjectTimeSeen.Exists(obsKey) Then
                duplicateCount = duplicateCount + 1
            Else
                subjectTimeSeen.Add obsKey, True
            End If
        End If
    Next r

    If groupOrder.Count = 0 Or timeOrder.Count = 0 Then Err.Raise vbObjectError + 5, , _
        "No complete numeric observations were found."

    Application.ScreenUpdating = False
    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = UniqueLongSheetName(wb, "RM Long Chart")

    wsOut.Cells(1, 1).Value = CStr(data(1, timeCol))
    outMeanStart = 2
    outSemStart = outMeanStart + groupOrder.Count
    outNStart = outSemStart + groupOrder.Count

    For j = 1 To groupOrder.Count
        wsOut.Cells(1, outMeanStart + j - 1).Value = CStr(groupOrder(j)) & " Mean"
        wsOut.Cells(1, outSemStart + j - 1).Value = CStr(groupOrder(j)) & " SEM"
        wsOut.Cells(1, outNStart + j - 1).Value = CStr(groupOrder(j)) & " N"
    Next j

    For i = 1 To timeOrder.Count
        wsOut.Cells(i + 1, 1).Value = CStr(timeOrder(i))
    Next i

    yMin = 1E+308
    yMax = -1E+308

    For j = 1 To groupOrder.Count
        groupName = CStr(groupOrder(j))
        For i = 1 To timeOrder.Count
            timeName = CStr(timeOrder(i))
            Set vals = New Collection

            For r = 2 To nRows
                If isGrouped Then
                    If Trim$(CStr(data(r, groupCol))) <> groupName Then GoTo ContinueRow
                End If
                If Trim$(CStr(data(r, timeCol))) = timeName And _
                   IsNumeric(data(r, valueCol)) And Len(CStr(data(r, valueCol))) > 0 And _
                   Len(Trim$(CStr(data(r, subjectCol)))) > 0 Then
                    vals.Add CDbl(data(r, valueCol))
                End If
ContinueRow:
            Next r

            validCount = vals.Count
            If validCount > 0 Then
                meanVal = LongCollectionMean(vals)
                sdVal = LongCollectionSampleSD(vals, meanVal)
                If validCount > 1 Then semVal = sdVal / Sqr(validCount) Else semVal = 0

                wsOut.Cells(i + 1, outMeanStart + j - 1).Value = meanVal
                wsOut.Cells(i + 1, outSemStart + j - 1).Value = semVal
                wsOut.Cells(i + 1, outNStart + j - 1).Value = validCount

                candidate = meanVal - semVal: If candidate < yMin Then yMin = candidate
                candidate = meanVal + semVal: If candidate > yMax Then yMax = candidate
            Else
                wsOut.Cells(i + 1, outMeanStart + j - 1).Formula = "=NA()"
                wsOut.Cells(i + 1, outSemStart + j - 1).Value = 0
                wsOut.Cells(i + 1, outNStart + j - 1).Value = 0
            End If
        Next i
    Next j

    Set timeRange = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(timeOrder.Count + 1, 1))
    Set chartObj = wsOut.ChartObjects.Add( _
        Left:=wsOut.Cells(2, outNStart + groupOrder.Count + 2).Left, _
        Top:=wsOut.Cells(2, outNStart + groupOrder.Count + 2).Top, _
        Width:=820, Height:=480)
    Set ch = chartObj.Chart
    ch.ChartType = xlLineMarkers
    ch.HasTitle = True
    ch.ChartTitle.Text = "Repeated-measures mean +/- SEM"

    colors = Array(RGB(91, 155, 213), RGB(237, 125, 49), RGB(112, 173, 71), _
                   RGB(165, 165, 165), RGB(255, 192, 0), RGB(68, 114, 196), _
                   RGB(153, 102, 204), RGB(0, 176, 240))

    For j = 1 To groupOrder.Count
        Set meanRange = wsOut.Range(wsOut.Cells(2, outMeanStart + j - 1), _
                                    wsOut.Cells(timeOrder.Count + 1, outMeanStart + j - 1))
        Set semRange = wsOut.Range(wsOut.Cells(2, outSemStart + j - 1), _
                                   wsOut.Cells(timeOrder.Count + 1, outSemStart + j - 1))
        Set s = ch.SeriesCollection.NewSeries
        s.Name = CStr(groupOrder(j))
        s.XValues = timeRange
        s.Values = meanRange
        clr = colors((j - 1) Mod (UBound(colors) + 1))
        s.Format.Line.ForeColor.RGB = clr
        s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 7
        s.MarkerForegroundColor = clr
        s.MarkerBackgroundColor = RGB(255, 255, 255)
        s.ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, _
                   Amount:=semRange, MinusValues:=semRange
    Next j

    ch.HasLegend = (groupOrder.Count > 1)
    If ch.HasLegend Then ch.Legend.Position = xlLegendPositionBottom

    With ch.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = CStr(data(1, timeCol))
        .TickLabelSpacing = 1
    End With
    With ch.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = CStr(data(1, valueCol))
    End With

    If yMax > -1E+307 And yMin < 1E+307 Then
        If yMax <= yMin Then
            If yMax = 0 Then yPad = 1 Else yPad = Abs(yMax) * 0.1
        Else
            yPad = (yMax - yMin) * 0.08
        End If
        If yMin >= 0 Then yMin = 0 Else yMin = yMin - yPad
        yMax = yMax + yPad
        ch.Axes(xlValue, xlPrimary).MinimumScale = yMin
        ch.Axes(xlValue, xlPrimary).MaximumScale = yMax
    End If

    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, outNStart + groupOrder.Count - 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    wsOut.Columns.AutoFit

    wsOut.Range("A" & timeOrder.Count + 4).Value = _
        "SEM = sample SD / SQRT(N); N is calculated separately for every group and time point."
    wsOut.Range("A" & timeOrder.Count + 4).Font.Italic = True

    If duplicateCount > 0 Then
        wsOut.Range("A" & timeOrder.Count + 5).Value = _
            "CAUTION: " & duplicateCount & " additional numeric row(s) shared the same Subject ID, Group, and Time. " & _
            "They were treated as separate observations. Average technical replicates first if appropriate."
        wsOut.Range("A" & timeOrder.Count + 5).Interior.Color = RGB(255, 235, 156)
        wsOut.Range("A" & timeOrder.Count + 5).WrapText = True
    End If

    wsOut.Activate
    chartObj.Activate
    Application.ScreenUpdating = True
    MsgBox "Long-format repeated-measures mean +/- SEM graph created on '" & wsOut.Name & "'.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Could not create the long-format repeated-measures graph: " & Err.Description, vbExclamation
End Sub

Private Function LongCollectionMean(ByVal values As Collection) As Double
    Dim x As Variant, total As Double
    For Each x In values: total = total + CDbl(x): Next x
    LongCollectionMean = total / values.Count
End Function

Private Function LongCollectionSampleSD(ByVal values As Collection, ByVal avg As Double) As Double
    Dim x As Variant, ss As Double
    If values.Count < 2 Then LongCollectionSampleSD = 0: Exit Function
    For Each x In values: ss = ss + (CDbl(x) - avg) ^ 2: Next x
    LongCollectionSampleSD = Sqr(ss / (values.Count - 1))
End Function

Private Function UniqueLongSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, n As Long
    candidate = baseName
    Do While LongSheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(baseName, 25) & " " & n
    Loop
    UniqueLongSheetName = candidate
End Function

Private Function LongSheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    LongSheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
