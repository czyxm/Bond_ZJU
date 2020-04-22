Attribute VB_Name = "Ä£¿é1"
Option Explicit
'@Brief: determine whether an item is in the range or not
Function isIn(item As Variant, container As Range) As Boolean
    isIn = Not container.Find(item, LookIn:=xlValues) Is Nothing
End Function
'@Brief: load one csv file from the filefolder and write data to out sheet
'@Return: the redundant date
Function loadFile(path As String, outSheet As Worksheet) As String
    Dim inWb As Workbook
    Set inWb = GetObject(path)
    Dim items As Variant, row As Integer, col As Integer, start As Integer
    start = outSheet.UsedRange.Rows.Count
    For row = 1 To inWb.Sheets(1).UsedRange.Rows.Count
        items = Split(inWb.Sheets(1).Cells(row, 1).Value, "|")
        If isIn(items(0), outSheet.Range("a1:a" & start)) Then
            loadFile = items(0)
            Exit Function
        Else
            For col = 0 To UBound(items)
                'Add Str(20) to the year values of date items
                If InStr(items(col), "-") > 0 Then
                    items(col) = Mid(items(col), 1, InStrRev(items(col), "-")) & "20" & Mid(items(col), InStrRev(items(col), "-") + 1)
                End If
                outSheet.Cells(row + start, col + 1).Value = items(col)
            Next
        End If
    Next
    loadFile = ""
End Function
'@Brief: load and deal with files from selected items in the folefolder
Sub loadFiles(sheet As Worksheet, name As String)
    Dim index As Integer, files As String, existed As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlw;*.csv"
        .Show
        
        For index = 1 To .SelectedItems.Count
            files = files & Chr(10) & index & ":: " & .SelectedItems(index)
            existed = existed & Chr(10) & loadFile(.SelectedItems(index), sheet)
        Next
        'Some prompts
        If .SelectedItems.Count > 0 Then
            MsgBox "Selected: " & files, vbOKOnly + vbInformation, name
        Else
            MsgBox "No file is selected.", vbOKOnly + vbInformation, name
        End If
        If Len(Replace(existed, Chr(10), "")) > 0 Then
            MsgBox "Existed date: " & existed, vbOKOnly + vbExclamation, name
        End If
    End With
End Sub
'@Brief: linear interpolation function for fetching yield
Function linearInterpolate(value1 As Double, value2 As Double, rate As Double) As Double
    linearInterpolate = value1 * (1 - rate) + value2 * rate
End Function
'@Brief: fetch yild form sheet "Yield.Curve"
'@Return: interpolated yield value
Function fetchYield(settlementDate As Date, couponDate As Date) As Double
    Dim row As Range
    'Find the corresponded row with the same date
    Set row = Worksheets("Yield.Curve").Range("a:a").Find(settlementDate, LookIn:=xlFormulas, lookat:=xlWhole)
    If row Is Nothing Then
        MsgBox "[Error]No such a date in the Sheet Yield.Curve " _
        & Chr(10) & "Date: " & settlementDate _
        , vbOKOnly + vbCritical, "Fetch Yield"
        Exit Function
    End If
    Dim duration
    Dim diff As Double
    Dim index As Integer
    'Maturity duration by month
    duration = Array(1, 3, 6, 12, 24, 36, 60, 84, 120, 240, 360)
    'Calculate the total maturity duration by day
    diff = DateDiff("m", settlementDate, couponDate)
    'Find the right position of maturity duration
    For index = LBound(duration) To UBound(duration)
        If duration(index) >= diff Then
            Exit For
        End If
    Next
    Dim startDate As Date, endDate As Date
    'Find the last date and the next date between which the maturity locates
    startDate = DateAdd("m", duration(index - 1), settlementDate)
    endDate = DateAdd("m", duration(index), settlementDate)
    'Calculate the duration portion
    diff = DateDiff("d", startDate, couponDate) / DateDiff("d", startDate, endDate)
    'Calculate yield through linear interpolation
    fetchYield = linearInterpolate(row.Offset(0, index), row.Offset(0, index + 1), diff) / 100
    'MsgBox "From " & settlementDate & " To " & couponDate & " Yield: " & fetchYield _
    '& Chr(10) & "C1: " & row.Offset(0, index) & " C2: " & row.Offset(0, index + 1)
End Function
'@Brief: calculate duration portion for later use
Function calPeriodRatio(settlementDate As Date, couponDate As Date, frequency As Integer, mode As Boolean) As Double
    Dim lastCouponDate As Date, nextCouponDate As Date
    lastCouponDate = CDate(Year(settlementDate) & "-" & Month(couponDate) & "-" & Day(couponDate))
    While (DateDiff("d", settlementDate, lastCouponDate) > 0)
        lastCouponDate = DateAdd("m", -12 / frequency, lastCouponDate)
    Wend
    nextCouponDate = DateAdd("m", 12 / frequency, lastCouponDate)
    If mode Then
        calPeriodRatio = DateDiff("d", settlementDate, nextCouponDate) / DateDiff("d", lastCouponDate, nextCouponDate)
    Else
        calPeriodRatio = DateDiff("d", lastCouponDate, settlementDate) / DateDiff("d", lastCouponDate, nextCouponDate)
    End If
End Function
'@Brief: calculate accrued interest
Function calAccruedInterest(coupon As Double, frequency As Integer, settlementDate As Date, couponDate As Date) As Double
    calAccruedInterest = coupon / frequency * calPeriodRatio(settlementDate, couponDate, frequency, False)
End Function
'@Brief: calculate a
Function calA(frequency As Integer, settlementDate As Date, couponDate As Date) As Double
    calA = calPeriodRatio(settlementDate, couponDate, frequency, True)
End Function
'@Brief: calculate dirty price and modified duration
'@Return: array of PV and MD
Function calDirtyPriceAndModifiedDuration(coupon As Double, frequency As Integer, settlementDate As Date, couponDate As Date) As Variant
    Dim yield As Double, cashflow As Double, a  As Double, PV As Double, MD As Double
    Dim i As Integer, m As Integer, D1 As Double, item As Double
    'Calculate a
    a = calA(frequency, settlementDate, couponDate)
    'Fetch yield
    yield = fetchYield(settlementDate, couponDate)
    'Calculate number of complete coupon period
    m = Int(DateDiff("m", settlementDate, couponDate) / 12 * frequency)
    'Cashflow in the beginning
    cashflow = coupon / frequency
    'Factor for later use
    D1 = 1 / (1 + yield / frequency)
    PV = 0
    MD = 0
    For i = 0 To m
        item = D1 ^ (a + i)
        PV = PV + item
        MD = MD + (a + i) * item
    Next
    PV = coupon / frequency * PV + 100 * item
    MD = (coupon / frequency * MD + 100 * (a + m) * item) / PV / frequency * D1
    calDirtyPriceAndModifiedDuration = Array(PV, MD)
End Function










