Attribute VB_Name = "PTVN_DashboardCharts"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"

Public Sub Auto_Open()
    RefreshPTVNDashboardCharts
End Sub

Public Sub RefreshPTVNDashboardCharts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.CalculateFullRebuild

    ClearDashboardCharts ws
    AddAttendanceStatusChart ws
    AddMonthlyTrendChart ws
    AddDepartmentPerformanceChart ws
    AddAbsenceBreakdownChart ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub ClearDashboardCharts(ByVal ws As Worksheet)
    Dim chartObject As ChartObject
    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
End Sub

Private Sub AddAttendanceStatusChart(ByVal ws As Worksheet)
    Dim chartObject As ChartObject
    Set chartObject = AddChartBox(ws, "C9:I27")

    With chartObject.Chart
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.Text = "Pass vs Below Target"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Employees"
        .SeriesCollection(1).Values = ws.Range("AA14:AA15")
        .SeriesCollection(1).XValues = ws.Range("Z14:Z15")
        .SeriesCollection(1).HasDataLabels = False
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .ChartGroups(1).DoughnutHoleSize = 58
        FormatCleanChart chartObject.Chart
        AddDonutCenterText chartObject, ws
    End With
End Sub

Private Sub AddMonthlyTrendChart(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws, "AC", 14)
    Dim pointCount As Long
    pointCount = lastRow - 13

    Dim chartObject As ChartObject
    Set chartObject = AddChartBox(ws, "J9:P27")

    With chartObject.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "PTV Overall Trend"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "PTV Trend"
        .SeriesCollection(1).Values = ws.Range("AD14:AD" & lastRow)
        .SeriesCollection(1).XValues = ws.Range("AC14:AC" & lastRow)
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(47, 85, 151)
        .SeriesCollection(1).Format.Line.Weight = 2.25
        .SeriesCollection(1).HasDataLabels = False
        If pointCount > 0 Then
            ClearPointLabels .SeriesCollection(1)
            .SeriesCollection(1).Points(pointCount).ApplyDataLabels
            .SeriesCollection(1).Points(pointCount).DataLabel.Text = _
                ShortMonthName(ws.Range("AC" & lastRow).Value) & " " & Format(ws.Range("AD" & lastRow).Value, "0%")
            .SeriesCollection(1).Points(pointCount).DataLabel.Position = xlLabelPositionRight
        End If

        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Target 60%"
        .SeriesCollection(2).Values = ws.Range("AE14:AE" & lastRow)
        .SeriesCollection(2).XValues = ws.Range("AC14:AC" & lastRow)
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(204, 0, 0)
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(2).Format.Line.Weight = 1.5
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone

        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MajorUnit = 0.2
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        FormatCleanChart chartObject.Chart
    End With
End Sub

Private Sub AddDepartmentPerformanceChart(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = LastUsedRow(ws, "W", 14)

    Dim chartObject As ChartObject
    Set chartObject = AddChartBox(ws, "C29:I49")

    With chartObject.Chart
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Average Attendance Rate by Department (%)"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Attendance"
        .SeriesCollection(1).Values = ws.Range("X14:X" & lastRow)
        .SeriesCollection(1).XValues = ws.Range("W14:W" & lastRow)
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.NumberFormat = "0%"
        .SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(183, 215, 168)
        ApplyDepartmentPerformanceLabels .SeriesCollection(1), ws, 14, lastRow

        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MajorUnit = 0.2
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = False
        FormatCleanChart chartObject.Chart
        AddVerticalTargetLine chartObject, 0.6
    End With
End Sub

Private Sub AddAbsenceBreakdownChart(ByVal ws As Worksheet)
    If WorksheetFunction.Sum(ws.Range("AG41:AG45")) = 0 Then
        ws.Range("J36:M43").ClearContents
        ws.Range("J38").Value = "No absence category data available"
        ws.Range("J38").Font.Bold = True
        ws.Range("J38").Font.Color = RGB(47, 85, 151)
        ws.Range("J39").Value = "Total absence days are calculated from attendance gap."
        ws.Range("J39").Font.Color = RGB(100, 116, 139)
        Exit Sub
    End If

    Dim chartObject As ChartObject
    Set chartObject = AddChartBox(ws, "J35:M45")

    With chartObject.Chart
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Absence Breakdown"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Days"
        .SeriesCollection(1).Values = ws.Range("AG40:AG45")
        .SeriesCollection(1).XValues = ws.Range("AF40:AF45")
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.NumberFormat = "0.0"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(125, 167, 217)
        .HasLegend = False
        FormatCleanChart chartObject.Chart
    End With
End Sub

Private Function AddChartBox(ByVal ws As Worksheet, ByVal address As String) As ChartObject
    Dim target As Range
    Set target = ws.Range(address)
    Set AddChartBox = ws.ChartObjects.Add(target.Left + 6, target.Top + 6, target.Width - 12, target.Height - 12)
End Function

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal columnLetter As String, ByVal minRow As Long) As Long
    Dim rowNumber As Long
    rowNumber = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
    If rowNumber < minRow Then
        rowNumber = minRow
    End If
    LastUsedRow = rowNumber
End Function

Private Sub ClearPointLabels(ByVal series As Series)
    Dim pointIndex As Long
    On Error Resume Next
    For pointIndex = 1 To series.Points.Count
        series.Points(pointIndex).HasDataLabel = False
    Next pointIndex
    On Error GoTo 0
End Sub

Private Sub ApplyDepartmentPerformanceLabels(ByVal series As Series, ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim pointIndex As Long
    Dim sourceRow As Long
    For pointIndex = 1 To series.Points.Count
        sourceRow = firstRow + pointIndex - 1
        series.Points(pointIndex).ApplyDataLabels
        series.Points(pointIndex).DataLabel.Text = ws.Range("W" & sourceRow).Value & "  " & Format(ws.Range("X" & sourceRow).Value, "0%")
        series.Points(pointIndex).DataLabel.Position = xlLabelPositionOutsideEnd
        series.Points(pointIndex).DataLabel.Font.Size = 8
    Next pointIndex
End Sub

Private Sub FormatCleanChart(ByVal chart As Chart)
    On Error Resume Next
    chart.ChartArea.Format.Fill.Visible = msoFalse
    chart.ChartArea.Format.Line.Visible = msoFalse
    chart.PlotArea.Format.Fill.Visible = msoFalse
    chart.PlotArea.Format.Line.Visible = msoFalse
    chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
    chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(15, 23, 42)
    chart.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(229, 231, 235)
    chart.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
    On Error GoTo 0
End Sub

Private Sub AddDonutCenterText(ByVal chartObject As ChartObject, ByVal ws As Worksheet)
    Dim totalEmployees As Double
    totalEmployees = WorksheetFunction.Sum(ws.Range("AA14:AA15"))

    Dim centerText As Shape
    Set centerText = chartObject.Chart.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        chartObject.Chart.ChartArea.Width * 0.39, _
        chartObject.Chart.ChartArea.Height * 0.43, _
        chartObject.Chart.ChartArea.Width * 0.22, _
        chartObject.Chart.ChartArea.Height * 0.16)

    With centerText
        .TextFrame2.TextRange.Text = Format(totalEmployees, "0") & vbCrLf & "Employees"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(15, 23, 42)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
End Sub

Private Sub AddVerticalTargetLine(ByVal chartObject As ChartObject, ByVal targetValue As Double)
    Dim chart As Chart
    Set chart = chartObject.Chart

    Dim plotLeft As Double
    Dim plotTop As Double
    Dim plotWidth As Double
    Dim plotHeight As Double
    plotLeft = chart.PlotArea.InsideLeft
    plotTop = chart.PlotArea.InsideTop
    plotWidth = chart.PlotArea.InsideWidth
    plotHeight = chart.PlotArea.InsideHeight

    Dim xPosition As Double
    xPosition = plotLeft + (plotWidth * targetValue)

    Dim lineShape As Shape
    Set lineShape = chart.Shapes.AddLine(xPosition, plotTop, xPosition, plotTop + plotHeight)
    With lineShape.Line
        .ForeColor.RGB = RGB(204, 0, 0)
        .DashStyle = msoLineDash
        .Weight = 1.25
    End With

    Dim labelShape As Shape
    Set labelShape = chart.Shapes.AddTextbox(msoTextOrientationHorizontal, xPosition + 4, plotTop + 2, 70, 16)
    With labelShape
        .TextFrame2.TextRange.Text = "Target 60%"
        .TextFrame2.TextRange.Font.Size = 8
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(204, 0, 0)
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
End Sub

Private Function ShortMonthName(ByVal value As Variant) As String
    Dim textValue As String
    textValue = CStr(value)
    If Len(textValue) >= 3 Then
        ShortMonthName = Left(textValue, 3)
    Else
        ShortMonthName = textValue
    End If
End Function
