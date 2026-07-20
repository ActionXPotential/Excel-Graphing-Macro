Attribute VB_Name = "RepeatedMeasuresMeanSEM"
Option Explicit

' Creates a repeated-measures mean +/- SEM line graph from wide-format data.
'
' Ungrouped layout:
' Subject ID | Time 1 | Time 2 | ... | Time N
'
' Grouped layout:
' Subject ID | Group | Time 1 | Time 2 | ... | Time N
'
' The macro graphs group means only. Individual subject trajectories are not plotted.

Public Sub CreateRepeatedMeasuresMeanSEMChart()
    Dim src As Range, wb As Workbook, wsOut As Worksheet
    Dim data As Variant, groupedAnswer As VbMsgBoxResult
    Dim isGrouped As Boolean, firstTimeCol As Long
    Dim nRows As Long, nCols As Long, nTimes As Long
    Dim groups As Object, groupOrder As Collection
    Dim g As String, key As String, vals As Collection
    Dim r As Long, c As Long, i As Long, j As Long
    Dim outMeanStart As Long, outSemStart As Long, outNStart As Long
    Dim meanVal As Double, sdVal As Double, semVal As Double
    Dim chObj As ChartObject, ch As Chart, s As Series
    Dim meanRng As Range, semRng As Range, timeRng As Range
    Dim colors As Variant, clr As Long
    Dim yMin As Double, yMax As Double, candidate As Double, yPad As Double
    Dim validCount As Long

    On Error Resume Next
    Set src = Application.InputBox( _
        Prompt:="Select the complete repeated-measures table INCLUDING headers." & vbCrLf & vbCrLf & _
               "Ungrouped: Subject ID | Time 1 | Time 2 | ..." & vbCrLf & _
               "Grouped: Subject ID | Group | Time 1 | Time 2 | ...", _
        Title:="Repeated-measures mean +/- SEM line graph", Type:=8)
    On Error GoTo CleanFail

    If src Is Nothing Then Exit Sub
    If src.Areas.Count > 1 Then Err.Raise vbObjectError + 1, , "Select one contiguous range."

    nRows = src.Rows.Count
    nCols = src.Columns.Count
    If nRows < 3 Then Err.Raise vbObjectError + 2, , "The selection needs a header and at least two subject rows."
    If nCols < 3 Then Err.Raise vbObjectError + 3, , "Select a Subject ID column and at least two time-point columns."

    groupedAnswer = MsgBox( _
        "Does column 2 contain a grouping variable (for example, Treatment or Genotype)?" & vbCrLf & vbCrLf & _
        "Yes: Subject ID | Group | Time 1 | Time 2 | ..." & vbCrLf & _
        "No:  Subject ID | Time 1 | Time 2 | ...", _
        vbYesNoCancel + vbQuestion, "Grouped repeated-measures data?")
    If groupedAnswer = vbCancel Then Exit Sub

    isGrouped = (groupedAnswer = vbYes)
    If isGrouped Then firstTimeCol = 3 Else firstTimeCol = 2
    nTimes = nCols - firstTimeCol + 1
    If nTimes < 2 Then Err.Raise vbObjectError + 4, , "At least two time-point columns are required."

    Set wb = src.Worksheet.Parent
    data = src.Value2
    Set groups = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection

    ' Preserve the first-appearance order of groups.
    If isGrouped Then
        For r = 2 To nRows
            g = Trim$(CStr(data(r, 2)))
            If Len(g) > 0 Then
                If Not groups.Exists(g) Then
                    groups.Add g, True
                    groupOrder.Add g
                End If
            End If
        Next r
    Else
        groups.Add "Mean", True
        groupOrder.Add "Mean"
    End If

    If groupOrder.Count = 0 Then Err.Raise vbObjectError + 5, , "No valid groups were found in column 2."

    Application.ScreenUpdating = False
    Set wsOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsOut.Name = UniqueSheetName(wb, "Repeated Measures Chart")

    ' Output matrix: time points in rows, one mean/SEM/N column per group.
    wsOut.Cells(1, 1).Value = "Time point"
    outMeanStart = 2
    outSemStart = outMeanStart + groupOrder.Count
    outNStart = outSemStart + groupOrder.Count

    For j = 1 To groupOrder.Count
        wsOut.Cells(1, outMeanStart + j - 1).Value = CStr(groupOrder(j)) & " Mean"
        wsOut.Cells(1, outSemStart + j - 1).Value = CStr(groupOrder(j)) & " SEM"
        wsOut.Cells(1, outNStart + j - 1).Value = CStr(groupOrder(j)) & " N"
    Next j

    For i = 1 To nTimes
        wsOut.Cells(i + 1, 1).Value = data(1, firstTimeCol + i - 1)
    Next i

    yMin = 1E+308
    yMax = -1E+308

    For j = 1 To groupOrder.Count
        g = CStr(groupOrder(j))
        For i = 1 To nTimes
            Set vals = New Collection
            c = firstTimeCol + i - 1

            For r = 2 To nRows
                If (Not isGrouped) Or Trim$(CStr(data(r, 2))) = g Then
                    If IsNumeric(data(r, c)) And Len(CStr(data(r, c))) > 0 Then
                        vals.Add CDbl(data(r, c))
                    End If
                End If
            Next r

            validCount = vals.Count
            If validCount > 0 Then
                meanVal = CollectionMean(vals)
                sdVal = CollectionSampleSD(vals, meanVal)
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

    Set timeRng = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(nTimes + 1, 1))
    Set chObj = wsOut.ChartObjects.Add( _
        Left:=wsOut.Cells(2, outNStart + groupOrder.Count + 2).Left, _
        Top:=wsOut.Cells(2, outNStart + groupOrder.Count + 2).Top, _
        Width:=820, Height:=480)
    Set ch = chObj.Chart
    ch.ChartType = xlLineMarkers
    ch.HasTitle = True
    ch.ChartTitle.Text = "Repeated-measures mean +/- SEM"

    colors = Array(RGB(91, 155, 213), RGB(237, 125, 49), RGB(112, 173, 71), _
                   RGB(165, 165, 165), RGB(255, 192, 0), RGB(68, 114, 196), _
                   RGB(153, 102, 204), RGB(0, 176, 240))

    For j = 1 To groupOrder.Count
        Set meanRng = wsOut.Range(wsOut.Cells(2, outMeanStart + j - 1), wsOut.Cells(nTimes + 1, outMeanStart + j - 1))
        Set semRng = wsOut.Range(wsOut.Cells(2, outSemStart + j - 1), wsOut.Cells(nTimes + 1, outSemStart + j - 1))
        Set s = ch.SeriesCollection.NewSeries
        s.Name = CStr(groupOrder(j))
        s.XValues = timeRng
        s.Values = meanRng
        clr = colors((j - 1) Mod (UBound(colors) + 1))
        s.Format.Line.ForeColor.RGB = clr
        s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 7
        s.MarkerForegroundColor = clr
        s.MarkerBackgroundColor = RGB(255, 255, 255)
        s.ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=semRng, MinusValues:=semRng
    Next j

    ch.HasLegend = (groupOrder.Count > 1)
    If ch.HasLegend Then ch.Legend.Position = xlLegendPositionBottom

    With ch.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Time"
        .TickLabelSpacing = 1
    End With
    With ch.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Mean +/- SEM"
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
    wsOut.Range("A" & nTimes + 4).Value = "SEM = sample SD / SQRT(N); N is calculated separately at each time point."
    wsOut.Range("A" & nTimes + 4).Font.Italic = True

    wsOut.Activate
    chObj.Activate
    Application.ScreenUpdating = True
    MsgBox "Repeated-measures mean +/- SEM line graph created on '" & wsOut.Name & "'.", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Could not create repeated-measures chart: " & Err.Description, vbExclamation
End Sub

Private Function CollectionMean(ByVal values As Collection) As Double
    Dim x As Variant, total As Double
    For Each x In values: total = total + CDbl(x): Next x
    CollectionMean = total / values.Count
End Function

Private Function CollectionSampleSD(ByVal values As Collection, ByVal avg As Double) As Double
    Dim x As Variant, ss As Double
    If values.Count < 2 Then CollectionSampleSD = 0: Exit Function
    For Each x In values: ss = ss + (CDbl(x) - avg) ^ 2: Next x
    CollectionSampleSD = Sqr(ss / (values.Count - 1))
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, n As Long
    candidate = baseName
    Do While SheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(baseName, 25) & " " & n
    Loop
    UniqueSheetName = candidate
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
