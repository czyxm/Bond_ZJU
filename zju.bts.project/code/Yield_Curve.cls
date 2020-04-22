VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Load_Yield_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3. col A to col L will be populated with the data from yc.2012.MM.DD.csv
    loadFiles Worksheets("Yield.Curve"), "Load Yield"
End Sub
'@Brief: plot Yield.Curve
Private Sub Plot_Yield_Click()
    Dim chart As ChartObject, sheet As Worksheet
    Dim row As Integer, endRow As Integer
    Set sheet = Worksheets("Yield.Curve")
    endRow = sheet.UsedRange.Rows.Count
    Application.ScreenUpdating = False
    'Clear the old charts
    If sheet.ChartObjects.Count > 0 Then
        sheet.ChartObjects.Delete
    End If
    'Add a chart at [m5]
    Set chart = sheet.ChartObjects.Add(sheet.[m5].Left, sheet.[m5].Top, 300, 300)
    With chart.chart
        .HasTitle = True
        .ChartTitle.Text = "Yield Curve"
        .ChartTitle.Font.Size = 18
        .HasLegend = True
        .Legend.Font.Size = 8
        .Legend.Font.ColorIndex = 5
        .Legend.position = xlLegendPositionRight
        'Config the Y label
        With .Axes(xlValue, xlPrimary)
            .CrossesAt = .MinimumScale
            .TickLabels.Font.Size = 8
            .HasTitle = True
            .AxisTitle.Text = "Yield(%)"
            .AxisTitle.Characters.Font.Size = 12
            .AxisTitle.Orientation = xlUpward
        End With
        'Config the X label
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 8
            .TickLabels.Orientation = 50
            .HasTitle = True
            .AxisTitle.Text = "Maturity"
            .AxisTitle.Characters.Font.Size = 12
        End With
        For row = 2 To endRow
            With .SeriesCollection.NewSeries
                .Values = sheet.Range("b" & row & ":l" & row)
                .XValues = sheet.Range("b1:l1")
                .ChartType = xlLineMarkers
                .name = sheet.Range("A" & row).Value
                .MarkerSize = 3
                .Format.Line.Weight = 0.8
                .MarkerStyle = xlMarkerStyleTriangle
                .MarkerForegroundColor = RGB(171, 130, 255)
                .MarkerBackgroundColor = RGB(171, 130, 255)
                .HasDataLabels = False
            End With
        Next
    End With
End Sub
