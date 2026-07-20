Attribute VB_Name = "UnifiedOutcomeGraphGenerator"
Option Explicit

' Unified chart generator
' Input layouts:
'   One-way: Group | Outcome 1 | Outcome 2 | ...
'   Two-way: Factor A | Factor B | Outcome 1 | Outcome 2 | ...
' Creates mean +/- SEM bars with individual-value overlays.
Public Sub GenerateUnifiedOutcomeGraphs()
    Dim src As Range, d As Variant, ans As Variant
    Dim design As Long, layout As Long, clusterMode As Long, orient As Long
    Dim firstOutcome As Long, categoryCol As Long, seriesCol As Long
    Dim wb As Workbook, idx As Worksheet, ws As Worksheet
    Dim c As Long, made As Long, doPNG As VbMsgBoxResult, pngFolder As String

    On Error Resume Next
    Set src = Application.InputBox("Select the complete contiguous table INCLUDING headers." & vbCrLf & vbCrLf & _
        "One-way: Group | Outcome 1 | Outcome 2 | ..." & vbCrLf & _
        "Two-way: Factor A | Factor B | Outcome 1 | Outcome 2 | ...", _
        "Unified Outcome Graph Generator", Type:=8)
    On Error GoTo Fail
    If src Is Nothing Then Exit Sub
    If src.Areas.Count > 1 Or src.Rows.Count < 3 Or src.Columns.Count < 2 Then _
        Err.Raise vbObjectError + 1, , "Select one contiguous range with headers and at least two data rows."

    ans = Application.InputBox("Choose the experimental design:" & vbCrLf & _
        "1 = One-way" & vbCrLf & "2 = Two-way", "Design", "1", Type:=2)
    If ans = False Then Exit Sub
    If Trim$(CStr(ans)) <> "1" And Trim$(CStr(ans)) <> "2" Then Err.Raise vbObjectError + 2, , "Enter 1 or 2."
    design = CLng(ans)
    firstOutcome = IIf(design = 1, 2, 3)
    If src.Columns.Count < firstOutcome Then Err.Raise vbObjectError + 3, , "No outcome columns were selected."

    ans = Application.InputBox("Choose the output layout:" & vbCrLf & _
        "1 = All outcomes on ONE graph" & vbCrLf & _
        "2 = One SEPARATE graph per outcome", "Graph layout", "1", Type:=2)
    If ans = False Then Exit Sub
    If Trim$(CStr(ans)) <> "1" And Trim$(CStr(ans)) <> "2" Then Err.Raise vbObjectError + 4, , "Enter 1 or 2."
    layout = CLng(ans)

    If layout = 1 Then
        ans = Application.InputBox("How should the single graph be clustered?" & vbCrLf & _
            "1 = Experimental groups/factor levels on the X axis; outcomes are bar series" & vbCrLf & _
            "2 = Outcomes on the X axis; groups/factor combinations are bar series", _
            "Single-graph clustering", "1", Type:=2)
        If ans = False Then Exit Sub
        If Trim$(CStr(ans)) <> "1" And Trim$(CStr(ans)) <> "2" Then Err.Raise vbObjectError + 5, , "Enter 1 or 2."
        clusterMode = CLng(ans)
    Else
        clusterMode = 1
    End If

    orient = 1
    If design = 2 And (layout = 2 Or clusterMode = 1) Then
        ans = Application.InputBox("Choose the two-way orientation:" & vbCrLf & _
            "1 = Column 1 on X axis; Column 2 as grouping series" & vbCrLf & _
            "2 = Column 2 on X axis; Column 1 as grouping series", _
            "Factor orientation", "1", Type:=2)
        If ans = False Then Exit Sub
        If Trim$(CStr(ans)) <> "1" And Trim$(CStr(ans)) <> "2" Then Err.Raise vbObjectError + 6, , "Enter 1 or 2."
        orient = CLng(ans)
    End If
    categoryCol = IIf(orient = 1, 1, 2)
    seriesCol = IIf(orient = 1, 2, 1)

    doPNG = MsgBox("Also export each generated chart as PNG?", vbYesNo + vbQuestion, "PNG export")
    If doPNG = vbYes Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Choose the PNG export folder"
            If .Show <> -1 Then Exit Sub
            pngFolder = .SelectedItems(1)
        End With
    End If

    Set wb = src.Worksheet.Parent
    d = src.Value2
    Application.ScreenUpdating = False

    If layout = 1 Then
        If clusterMode = 1 Then
            Set ws = BuildGroupClustered(src, design, categoryCol, seriesCol, firstOutcome)
        Else
            Set ws = BuildOutcomeClustered(src, design, firstOutcome)
        End If
        made = 1
        If doPNG = vbYes Then ws.ChartObjects(1).Chart.Export pngFolder & Application.PathSeparator & SafeFile(CStr(ws.ChartObjects(1).Chart.ChartTitle.Text)) & ".png", "PNG"
    Else
        Set idx = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        idx.Name = UniqueSheet(wb, "Graph Index")
        idx.Range("A1:D1").Value = Array("Outcome", "Chart sheet", "Source column", "PNG path")
        StyleHeader idx.Range("A1:D1")
        For c = firstOutcome To UBound(d, 2)
            Set ws = BuildSeparateOutcome(src, design, categoryCol, seriesCol, c)
            made = made + 1
            idx.Cells(made + 1, 1).Value = CStr(d(1, c))
            idx.Cells(made + 1, 2).Value = ws.Name
            idx.Hyperlinks.Add idx.Cells(made + 1, 2), "", "'" & ws.Name & "'!A1", , ws.Name
            idx.Cells(made + 1, 3).Value = src.Columns(c).Address(False, False)
            If doPNG = vbYes Then
                idx.Cells(made + 1, 4).Value = pngFolder & Application.PathSeparator & SafeFile(CStr(d(1, c))) & ".png"
                ws.ChartObjects(1).Chart.Export CStr(idx.Cells(made + 1, 4).Value), "PNG"
            End If
        Next c
        idx.Columns.AutoFit
        idx.Activate
    End If

    Application.ScreenUpdating = True
    MsgBox made & " chart(s) created. Raw points are centered unless values overlap; both Y axes use the same data-derived limits.", vbInformation
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Could not generate graphs: " & Err.Description, vbExclamation
End Sub

Private Function BuildSeparateOutcome(src As Range, design As Long, categoryCol As Long, seriesCol As Long, outcomeCol As Long) As Worksheet
    Dim outcomes As Collection: Set outcomes = New Collection: outcomes.Add outcomeCol
    Set BuildSeparateOutcome = BuildCategoryChart(src, design, categoryCol, seriesCol, outcomes, CStr(src.Cells(1, outcomeCol).Value), False)
End Function

Private Function BuildGroupClustered(src As Range, design As Long, categoryCol As Long, seriesCol As Long, firstOutcome As Long) As Worksheet
    Dim outcomes As New Collection, c As Long
    For c = firstOutcome To src.Columns.Count: outcomes.Add c: Next c
    Set BuildGroupClustered = BuildCategoryChart(src, design, categoryCol, seriesCol, outcomes, "Multi-outcome chart clustered by group", True)
End Function

Private Function BuildCategoryChart(src As Range, design As Long, categoryCol As Long, seriesCol As Long, outcomes As Collection, titleText As String, combineOutcomes As Boolean) As Worksheet
    Dim d As Variant, wb As Workbook, ws As Worksheet, cats As Object, facs As Object, groups As Object
    Dim catOrder As New Collection, facOrder As New Collection, vals As Collection
    Dim r As Long, i As Long, j As Long, m As Long, p As Long, sIdx As Long
    Dim cat As String, fac As String, outName As String, key As String, label As String
    Dim nCat As Long, nFac As Long, nSer As Long, meanStart As Long, semStart As Long, nStart As Long
    Dim meanV As Double, semV As Double, yMin As Double, yMax As Double, yPad As Double
    Dim pointHead As Long, pointRow As Long, barCenter As Double, slotW As Double
    Dim chObj As ChartObject, ch As Chart, s As Series, colors As Variant, clr As Long
    Dim x() As Double, y() As Double, pointCount As Long, headers As Range

    d = src.Value2: Set wb = src.Worksheet.Parent
    Set cats = CreateObject("Scripting.Dictionary"): Set facs = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(d, 1)
        cat = Trim$(CStr(d(r, categoryCol)))
        fac = IIf(design = 1, "All", Trim$(CStr(d(r, seriesCol))))
        If Len(cat) > 0 And Len(fac) > 0 Then
            If Not cats.Exists(cat) Then cats.Add cat, True: catOrder.Add cat
            If Not facs.Exists(fac) Then facs.Add fac, True: facOrder.Add fac
            For m = 1 To outcomes.Count
                If IsNumeric(d(r, CLng(outcomes(m)))) And Len(CStr(d(r, CLng(outcomes(m))))) > 0 Then
                    outName = CStr(d(1, CLng(outcomes(m))))
                    key = MakeKey(cat, fac, outName)
                    If Not groups.Exists(key) Then Set vals = New Collection: groups.Add key, vals
                    groups(key).Add CDbl(d(r, CLng(outcomes(m))))
                End If
            Next m
        End If
    Next r
    nCat = catOrder.Count: nFac = facOrder.Count: nSer = nFac * outcomes.Count
    If nCat = 0 Or nSer = 0 Then Err.Raise vbObjectError + 20, , "No numeric observations were found."

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = UniqueSheet(wb, SafeSheet(Left$(titleText, 25)))
    meanStart = 2: semStart = meanStart + nSer: nStart = semStart + nSer
    ws.Cells(1, 1).Value = CStr(d(1, categoryCol))
    For i = 1 To nCat: ws.Cells(i + 1, 1).Value = CStr(catOrder(i)): Next i
    sIdx = 0
    For m = 1 To outcomes.Count
        outName = CStr(d(1, CLng(outcomes(m))))
        For j = 1 To nFac
            sIdx = sIdx + 1: fac = CStr(facOrder(j))
            If combineOutcomes Then
                label = IIf(design = 1, outName, fac & " - " & outName)
            Else
                label = IIf(design = 1, outName, fac)
            End If
            ws.Cells(1, meanStart + sIdx - 1) = label & " Mean"
            ws.Cells(1, semStart + sIdx - 1) = label & " SEM"
            ws.Cells(1, nStart + sIdx - 1) = label & " N"
        Next j
    Next m

    yMin = 1E+308: yMax = -1E+308: sIdx = 0
    For m = 1 To outcomes.Count
        outName = CStr(d(1, CLng(outcomes(m))))
        For j = 1 To nFac
            sIdx = sIdx + 1: fac = CStr(facOrder(j))
            For i = 1 To nCat
                key = MakeKey(CStr(catOrder(i)), fac, outName)
                If groups.Exists(key) Then
                    Set vals = groups(key): meanV = ColMean(vals): semV = ColSEM(vals, meanV)
                    ws.Cells(i + 1, meanStart + sIdx - 1) = meanV
                    ws.Cells(i + 1, semStart + sIdx - 1) = semV
                    ws.Cells(i + 1, nStart + sIdx - 1) = vals.Count
                    UpdateBounds yMin, yMax, meanV - semV: UpdateBounds yMin, yMax, meanV + semV
                Else
                    ws.Cells(i + 1, meanStart + sIdx - 1).Formula = "=NA()"
                End If
            Next i
        Next j
    Next m

    pointHead = nCat + 4: pointRow = pointHead + 1
    ws.Range(ws.Cells(pointHead, 1), ws.Cells(pointHead, 4)).Value = Array("Series", "Point X", "Point Y", "Bar key")
    slotW = 1 / (nSer + 0.55): sIdx = 0
    For m = 1 To outcomes.Count
        outName = CStr(d(1, CLng(outcomes(m))))
        For j = 1 To nFac
            sIdx = sIdx + 1: fac = CStr(facOrder(j))
            If combineOutcomes Then label = IIf(design = 1, outName, fac & " - " & outName) Else label = IIf(design = 1, outName, fac)
            For i = 1 To nCat
                key = MakeKey(CStr(catOrder(i)), fac, outName)
                If groups.Exists(key) Then
                    Set vals = groups(key): barCenter = i + (sIdx - (nSer + 1) / 2) * slotW
                    For p = 1 To vals.Count
                        ws.Cells(pointRow, 1) = label
                        ws.Cells(pointRow, 2) = barCenter + CollisionOffset(vals, p, slotW, yMin, yMax)
                        ws.Cells(pointRow, 3) = CDbl(vals(p))
                        ws.Cells(pointRow, 4) = key
                        UpdateBounds yMin, yMax, CDbl(vals(p))
                        pointRow = pointRow + 1
                    Next p
                End If
            Next i
        Next j
    Next m

    Set chObj = ws.ChartObjects.Add(80, 40, 900, 520): Set ch = chObj.Chart
    ch.ChartType = xlColumnClustered: ch.HasTitle = True: ch.ChartTitle.Text = titleText
    colors = Palette()
    For sIdx = 1 To nSer
        Set s = ch.SeriesCollection.NewSeries
        s.Name = Replace(CStr(ws.Cells(1, meanStart + sIdx - 1).Value), " Mean", "")
        s.XValues = ws.Range(ws.Cells(2, 1), ws.Cells(nCat + 1, 1))
        s.Values = ws.Range(ws.Cells(2, meanStart + sIdx - 1), ws.Cells(nCat + 1, meanStart + sIdx - 1))
        clr = colors((sIdx - 1) Mod (UBound(colors) + 1))
        s.Format.Fill.ForeColor.RGB = clr: s.Format.Fill.Transparency = 0.2: s.Format.Line.ForeColor.RGB = clr
        s.ErrorBar xlY, xlBoth, xlCustom, ws.Range(ws.Cells(2, semStart + sIdx - 1), ws.Cells(nCat + 1, semStart + sIdx - 1)), ws.Range(ws.Cells(2, semStart + sIdx - 1), ws.Cells(nCat + 1, semStart + sIdx - 1))
    Next sIdx
    ch.ChartGroups(1).GapWidth = 55
    AddPointSeries ch, ws, pointHead, pointRow, nSer, meanStart, colors
    SetAxes ch, nCat, yMin, yMax
    If nSer = 1 Then ch.HasLegend = False Else TrimPointLegend ch, nSer
    StyleHeader ws.Range(ws.Cells(1, 1), ws.Cells(1, nStart + nSer - 1)): StyleHeader ws.Range(ws.Cells(pointHead, 1), ws.Cells(pointHead, 4))
    ws.Columns.AutoFit
    Set BuildCategoryChart = ws
End Function

Private Function BuildOutcomeClustered(src As Range, design As Long, firstOutcome As Long) As Worksheet
    Dim d As Variant, wb As Workbook, ws As Worksheet, seen As Object, groups As Object, order As New Collection, vals As Collection
    Dim r As Long, i As Long, j As Long, p As Long, nOut As Long, nSer As Long, outName As String, serName As String, key As String
    Dim meanStart As Long, semStart As Long, nStart As Long, pointHead As Long, pointRow As Long
    Dim meanV As Double, semV As Double, yMin As Double, yMax As Double, slotW As Double, center As Double
    Dim chObj As ChartObject, ch As Chart, s As Series, colors As Variant, clr As Long
    d = src.Value2: Set wb = src.Worksheet.Parent: nOut = UBound(d, 2) - firstOutcome + 1
    Set seen = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(d, 1)
        If design = 1 Then serName = Trim$(CStr(d(r, 1))) Else serName = Trim$(CStr(d(r, 1))) & " | " & Trim$(CStr(d(r, 2)))
        If Len(serName) > 0 Then
            If Not seen.Exists(serName) Then seen.Add serName, True: order.Add serName
            For i = 1 To nOut
                If IsNumeric(d(r, firstOutcome + i - 1)) And Len(CStr(d(r, firstOutcome + i - 1))) > 0 Then
                    outName = CStr(d(1, firstOutcome + i - 1)): key = MakeKey(outName, serName, "")
                    If Not groups.Exists(key) Then Set vals = New Collection: groups.Add key, vals
                    groups(key).Add CDbl(d(r, firstOutcome + i - 1))
                End If
            Next i
        End If
    Next r
    nSer = order.Count: If nSer = 0 Then Err.Raise vbObjectError + 30, , "No valid groups were found."
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = UniqueSheet(wb, "Outcome Cluster Chart")
    meanStart = 2: semStart = meanStart + nSer: nStart = semStart + nSer
    ws.Cells(1, 1) = "Outcome"
    For j = 1 To nSer: ws.Cells(1, meanStart + j - 1) = order(j) & " Mean": ws.Cells(1, semStart + j - 1) = order(j) & " SEM": ws.Cells(1, nStart + j - 1) = order(j) & " N": Next j
    For i = 1 To nOut: ws.Cells(i + 1, 1) = CStr(d(1, firstOutcome + i - 1)): Next i
    yMin = 1E+308: yMax = -1E+308
    For j = 1 To nSer
        serName = CStr(order(j))
        For i = 1 To nOut
            outName = CStr(d(1, firstOutcome + i - 1)): key = MakeKey(outName, serName, "")
            If groups.Exists(key) Then
                Set vals = groups(key): meanV = ColMean(vals): semV = ColSEM(vals, meanV)
                ws.Cells(i + 1, meanStart + j - 1) = meanV: ws.Cells(i + 1, semStart + j - 1) = semV: ws.Cells(i + 1, nStart + j - 1) = vals.Count
                UpdateBounds yMin, yMax, meanV - semV: UpdateBounds yMin, yMax, meanV + semV
            Else: ws.Cells(i + 1, meanStart + j - 1).Formula = "=NA()"
            End If
        Next i
    Next j
    pointHead = nOut + 4: pointRow = pointHead + 1: ws.Range(ws.Cells(pointHead, 1), ws.Cells(pointHead, 4)).Value = Array("Series", "Point X", "Point Y", "Bar key")
    slotW = 1 / (nSer + 0.55)
    For j = 1 To nSer
        serName = CStr(order(j))
        For i = 1 To nOut
            outName = CStr(d(1, firstOutcome + i - 1)): key = MakeKey(outName, serName, "")
            If groups.Exists(key) Then
                Set vals = groups(key): center = i + (j - (nSer + 1) / 2) * slotW
                For p = 1 To vals.Count
                    ws.Cells(pointRow, 1) = serName: ws.Cells(pointRow, 2) = center + CollisionOffset(vals, p, slotW, yMin, yMax): ws.Cells(pointRow, 3) = CDbl(vals(p)): ws.Cells(pointRow, 4) = key
                    UpdateBounds yMin, yMax, CDbl(vals(p)): pointRow = pointRow + 1
                Next p
            End If
        Next i
    Next j
    Set chObj = ws.ChartObjects.Add(80, 40, 900, 520): Set ch = chObj.Chart: ch.ChartType = xlColumnClustered: ch.HasTitle = True: ch.ChartTitle.Text = "Outcomes clustered by group"
    colors = Palette()
    For j = 1 To nSer
        Set s = ch.SeriesCollection.NewSeries: s.Name = CStr(order(j)): s.XValues = ws.Range(ws.Cells(2, 1), ws.Cells(nOut + 1, 1)): s.Values = ws.Range(ws.Cells(2, meanStart + j - 1), ws.Cells(nOut + 1, meanStart + j - 1))
        clr = colors((j - 1) Mod (UBound(colors) + 1)): s.Format.Fill.ForeColor.RGB = clr: s.Format.Fill.Transparency = 0.2: s.Format.Line.ForeColor.RGB = clr
        s.ErrorBar xlY, xlBoth, xlCustom, ws.Range(ws.Cells(2, semStart + j - 1), ws.Cells(nOut + 1, semStart + j - 1)), ws.Range(ws.Cells(2, semStart + j - 1), ws.Cells(nOut + 1, semStart + j - 1))
    Next j
    ch.ChartGroups(1).GapWidth = 55: AddPointSeries ch, ws, pointHead, pointRow, nSer, meanStart, colors: SetAxes ch, nOut, yMin, yMax: TrimPointLegend ch, nSer
    StyleHeader ws.Range(ws.Cells(1, 1), ws.Cells(1, nStart + nSer - 1)): StyleHeader ws.Range(ws.Cells(pointHead, 1), ws.Cells(pointHead, 4)): ws.Columns.AutoFit
    Set BuildOutcomeClustered = ws
End Function

Private Sub AddPointSeries(ch As Chart, ws As Worksheet, pointHead As Long, pointRow As Long, nSer As Long, meanStart As Long, colors As Variant)
    Dim idx As Long, r As Long, p As Long, n As Long, label As String, x() As Double, y() As Double, s As Series
    For idx = 1 To nSer
        label = Replace(CStr(ws.Cells(1, meanStart + idx - 1).Value), " Mean", "")
        n = Application.WorksheetFunction.CountIf(ws.Range(ws.Cells(pointHead + 1, 1), ws.Cells(pointRow - 1, 1)), label)
        If n > 0 Then
            ReDim x(1 To n): ReDim y(1 To n): p = 0
            For r = pointHead + 1 To pointRow - 1
                If CStr(ws.Cells(r, 1).Value) = label Then p = p + 1: x(p) = CDbl(ws.Cells(r, 2).Value): y(p) = CDbl(ws.Cells(r, 3).Value)
            Next r
            Set s = ch.SeriesCollection.NewSeries: s.Name = label & " values": s.ChartType = xlXYScatter: s.AxisGroup = xlSecondary: s.XValues = x: s.Values = y
            s.MarkerStyle = xlMarkerStyleCircle: s.MarkerSize = 5: s.MarkerForegroundColor = RGB(35, 35, 35): s.MarkerBackgroundColor = colors((idx - 1) Mod (UBound(colors) + 1)): s.Format.Line.Visible = msoFalse
        End If
    Next idx
End Sub

' Points stay exactly at the bar center unless another value in the same bar is
' equal or visually near-equal (within 0.6% of the current plotted Y range).
Private Function CollisionOffset(vals As Collection, p As Long, slotW As Double, yMin As Double, yMax As Double) As Double
    Dim q As Long, rank As Long, countNear As Long, tol As Double, spread As Double
    spread = yMax - yMin: If spread <= 0 Or yMin > 1E+300 Then spread = 1
    tol = spread * 0.006
    For q = 1 To vals.Count
        If Abs(CDbl(vals(q)) - CDbl(vals(p))) <= tol Then
            countNear = countNear + 1
            If q <= p Then rank = rank + 1
        End If
    Next q
    If countNear <= 1 Then CollisionOffset = 0 Else CollisionOffset = ((rank - 1) - (countNear - 1) / 2) * (0.56 * slotW / Application.Max(countNear - 1, 1))
End Function

Private Sub SetAxes(ch As Chart, nCat As Long, yMin As Double, yMax As Double)
    Dim pad As Double, lo As Double, hi As Double
    If yMin > 1E+300 Then yMin = 0: yMax = 1
    If yMax <= yMin Then pad = IIf(yMax = 0, 1, Abs(yMax) * 0.1) Else pad = (yMax - yMin) * 0.08
    lo = yMin - pad: hi = yMax + pad
    If yMin >= 0 Then lo = 0
    If yMax <= 0 Then hi = 0
    ch.HasAxis(xlCategory, xlSecondary) = True: ch.HasAxis(xlValue, xlSecondary) = True
    With ch.Axes(xlCategory, xlSecondary): .MinimumScale = 0.5: .MaximumScale = nCat + 0.5: .MajorUnit = 1: .TickLabelPosition = xlNone: On Error Resume Next: .Border.LineStyle = xlNone: On Error GoTo 0: End With
    With ch.Axes(xlValue, xlPrimary): .MinimumScale = lo: .MaximumScale = hi: .HasTitle = True: .AxisTitle.Text = "Outcome value": End With
    With ch.Axes(xlValue, xlSecondary): .MinimumScale = lo: .MaximumScale = hi: .TickLabelPosition = xlNone: On Error Resume Next: .Border.LineStyle = xlNone: On Error GoTo 0: End With
End Sub

Private Sub TrimPointLegend(ch As Chart, nBarSeries As Long)
    Dim i As Long: ch.HasLegend = True
    For i = ch.Legend.LegendEntries.Count To nBarSeries + 1 Step -1: ch.Legend.LegendEntries(i).Delete: Next i
    ch.Legend.Position = xlLegendPositionBottom
End Sub
Private Sub UpdateBounds(ByRef lo As Double, ByRef hi As Double, ByVal v As Double)
    If v < lo Then lo = v
    If v > hi Then hi = v
End Sub

Private Function ColMean(ByVal v As Collection) As Double
    Dim x As Variant
    For Each x In v
        ColMean = ColMean + CDbl(x)
    Next x
    ColMean = ColMean / v.Count
End Function

Private Function ColSEM(ByVal v As Collection, ByVal m As Double) As Double
    Dim x As Variant, ss As Double
    If v.Count < 2 Then
        ColSEM = 0
        Exit Function
    End If
    For Each x In v
        ss = ss + (CDbl(x) - m) ^ 2
    Next x
    ColSEM = Sqr(ss / (v.Count - 1)) / Sqr(v.Count)
End Function

Private Function MakeKey(ByVal a As String, ByVal b As String, ByVal c As String) As String
    MakeKey = a & ChrW(30) & b & ChrW(30) & c
End Function

Private Function Palette() As Variant
    Palette = Array(RGB(91, 155, 213), RGB(237, 125, 49), RGB(112, 173, 71), _
        RGB(165, 165, 165), RGB(255, 192, 0), RGB(68, 114, 196), _
        RGB(153, 102, 204), RGB(0, 176, 240), RGB(192, 80, 77), _
        RGB(75, 172, 198), RGB(128, 100, 162), RGB(155, 187, 89))
End Function

Private Sub StyleHeader(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
End Sub

Private Function SafeFile(ByVal s As String) As String
    Dim x As Variant
    For Each x In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        s = Replace(s, CStr(x), "_")
    Next x
    SafeFile = s
End Function

Private Function SafeSheet(ByVal s As String) As String
    Dim x As Variant
    For Each x In Array("\", "/", ":", "*", "?", "[", "]")
        s = Replace(s, CStr(x), "_")
    Next x
    SafeSheet = Left$(s, 31)
End Function

Private Function UniqueSheet(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim n As Long

    candidate = Left$(baseName, 31)
    If Len(candidate) = 0 Then candidate = "Chart"

    Do While SheetExists(wb, candidate)
        n = n + 1
        candidate = Left$(baseName, 27) & "_" & CStr(n)
    Loop

    UniqueSheet = candidate
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
