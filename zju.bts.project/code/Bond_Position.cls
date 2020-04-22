VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Load_Position_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3. col A to col E will be populated with the data from bond.postion..2012.MM.DD.csv
    loadFiles Worksheets("Bond.Position"), "Load Position"
End Sub

Private Sub Calculate_Plot_Click()
'Pls ensure the following functions
'1. pop up a window for user to select the file to load
'2. if the data of a particular date exists, a warning msg will be shown
'3. col F to col J will be populated
'4. Plot the whole book's Position BPV curve. The book includes 4 bond (A, B, C and D)
    Dim sheet As Worksheet
    Dim coupon As Double, position As Double
    Dim settlementDate As Date, couponDate As Date
    Dim row As Integer, col As Integer, PV_MD As Variant
    Set sheet = Worksheets("Bond.Position")
    For row = 2 To sheet.UsedRange.Rows.Count
        'Fetch information
        settlementDate = sheet.Range("a" & row).Value
        coupon = sheet.Range("c" & row).Value
        couponDate = sheet.Range("d" & row).Value
        position = sheet.Range("e" & row).Value
        'Calculate dirty price and modified duration
        PV_MD = calDirtyPriceAndModifiedDuration(coupon, 2, settlementDate, couponDate)
        sheet.Range("h" & row).Value = PV_MD(0)
        sheet.Range("i" & row).Value = PV_MD(1)
        'Calculate accrued interest
        sheet.Range("g" & row).Value = calAccruedInterest(coupon, 2, settlementDate, couponDate)
        'Clean price = dirty price - accrued interest
        sheet.Range("f" & row).Value = sheet.Range("h" & row).Value - sheet.Range("g" & row).Value
        'Calculate position basis point value
        sheet.Range("j" & row).Value = sheet.Range("e" & row).Value * PV_MD(0) / 100 * PV_MD(1) / 100
    Next
    
    'Plot Position.BPV.Curve
    plotPositionBPV sheet, Worksheets("Position.BPV.Curve")
End Sub
'@Brief: search the range for xValues and yValues for line chart
Private Sub searchXYRange(sheet As Worksheet, ByRef xDictionary As Variant, ByRef yDictionary As Variant)
    Dim i As Integer, n As Integer, name As String, item
    n = sheet.UsedRange.Rows.Count
    For i = 2 To n
        'B: name A:XValue J:yValue
        name = sheet.Range("b" & i).Value
        xDictionary(name) = xDictionary(name) & "a" & i & ","
        yDictionary(name) = yDictionary(name) & "j" & i & ","
        'MsgBox "Name: " & name & Chr(10) & "X: " & xDictionary(name) & Chr(10) & "Y: " & yDictionary(name)
    Next
    For Each item In xDictionary.keys
        xDictionary(item) = Left(xDictionary(item), Len(xDictionary(item)) - 1)
        yDictionary(item) = Left(yDictionary(item), Len(yDictionary(item)) - 1)
    Next
End Sub
'@Brief: plot Position.BPV.Curve
Private Sub plotPositionBPV(sheet As Worksheet, outSheet As Worksheet)
    Dim chart As ChartObject
    Application.ScreenUpdating = False
    'Clear the old charts
    If outSheet.ChartObjects.Count > 0 Then
        outSheet.ChartObjects.Delete
    End If
    'Add a chart at [a1]
    Set chart = outSheet.ChartObjects.Add(sheet.[a1].Left, sheet.[a1].Top, 400, 300)
    With chart.chart
        'Config the chart
        .HasTitle = True
        .ChartTitle.Text = "Position.BPV.Curve"
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
            .AxisTitle.Text = "Position.BPV"
            .AxisTitle.Characters.Font.Size = 12
            .AxisTitle.Orientation = xlUpward
        End With
        'Config the X label
        With .Axes(xlCategory)
            .TickLabels.Font.Size = 8
            .TickLabels.Orientation = 50
            .HasTitle = True
            .AxisTitle.Text = "Settlement Date"
            .AxisTitle.Characters.Font.Size = 12
        End With
        Dim xDictionary, yDictionary, item
        Set xDictionary = CreateObject("Scripting.Dictionary")
        Set yDictionary = CreateObject("Scripting.Dictionary")
        searchXYRange sheet, xDictionary, yDictionary
        For Each item In xDictionary.keys
            'MsgBox "Name: " & item & " xValue: " & xDictionary(item) & " yValue: " & yDictionary(item)
            With .SeriesCollection.NewSeries
                .Values = sheet.Range(yDictionary(item))
                .XValues = sheet.Range(xDictionary(item))
                .ChartType = xlLineMarkers
                .name = item
                .MarkerSize = 5
                .Format.Line.Weight = 1
                .MarkerStyle = xlMarkerStyleTriangle
                .MarkerForegroundColor = RGB(171, 130, 255)
                .MarkerBackgroundColor = RGB(171, 130, 255)
                .HasDataLabels = False
            End With
        Next
    End With
End Sub
