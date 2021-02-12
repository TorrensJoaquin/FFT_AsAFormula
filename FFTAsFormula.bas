Option Explicit
Function PerformAFFT(TimeInSecsAsRange As Range, DataAsRange As Range) As Variant
Dim Data() As Variant
Dim Time() As Variant
Data = DataAsRange.Value2
Set DataAsRange = Nothing
Time = TimeInSecsAsRange.Value2
Set TimeInSecsAsRange = Nothing

Dim n As Long, x As Long
Dim TFactor_N1 As Variant, TFactor_N2 As Variant, TimeLapse As Double

n = UBound(Data, 1)
Do Until 2 ^ x <= n And 2 ^ (x + 1) > n                                                                     'locates largest power of 2 from size of input array
    x = x + 1
Loop

n = n - (n - 2 ^ x)

TimeLapse = Abs(Time(n, 1) - Time(1, 1))

TFactor_N1 = WorksheetFunction.ImExp(WorksheetFunction.Complex(0, -2 * WorksheetFunction.Pi / (n / 1)))
TFactor_N2 = WorksheetFunction.ImExp(WorksheetFunction.Complex(0, -2 * WorksheetFunction.Pi / (n / 2)))

PerformAFFT = FFTFunction(Data, n, TFactor_N1, TFactor_N2, n / 2 - 1, TimeLapse)

End Function
Function PerformAFFTDateFormat(TimeInDaysAsRange As Range, DataAsRange As Range) As Variant
Dim Data() As Variant
Dim Time() As Variant
Data = DataAsRange.Value2
Set DataAsRange = Nothing
Time = TimeInDaysAsRange.Value2
Set TimeInDaysAsRange = Nothing

Dim n As Long, x As Long
Dim TFactor_N1 As Variant, TFactor_N2 As Variant, TimeLapse As Double

n = UBound(Data, 1)
Do Until 2 ^ x <= n And 2 ^ (x + 1) > n                                                                     'locates largest power of 2 from size of input array
    x = x + 1
Loop

n = n - (n - 2 ^ x)

TimeLapse = Abs(Time(n, 1) - Time(1, 1)) * 24 * 3600

TFactor_N1 = WorksheetFunction.ImExp(WorksheetFunction.Complex(0, -2 * WorksheetFunction.Pi / (n / 1)))
TFactor_N2 = WorksheetFunction.ImExp(WorksheetFunction.Complex(0, -2 * WorksheetFunction.Pi / (n / 2)))

PerformAFFT = FFTFunction(Data, n, TFactor_N1, TFactor_N2, n / 2 - 1, TimeLapse)

End Function
Private Function FFTFunction(Data As Variant, n As Long, TFactor_N1 As Variant, TFactor_N2 As Variant, NumberOfResults As Long, TimeLapse As Double) As Variant

Dim Result() As Variant
ReDim Result(1 To n / 2, 1 To 2)

Dim f_1() As Variant, f_2() As Variant
Dim i As Long, m As Long, k As Long
Dim G_1() As Variant, G_2() As Variant

ReDim f_1(0 To NumberOfResults)
ReDim f_2(0 To NumberOfResults)
ReDim G_1(0 To n / 1 - 1) As Variant
ReDim G_2(0 To n / 1 - 1) As Variant
ReDim X_k(0 To n - 1) As Variant

For i = 0 To NumberOfResults
    f_1(i) = Data(2 * i + 1, 1)
    f_2(i) = Data(2 * i + 2, 1)
Next i
For k = 0 To NumberOfResults
    For m = 0 To NumberOfResults
        G_1(m) = WorksheetFunction.ImProduct(WorksheetFunction.ImPower(TFactor_N2, k * m), f_1(m))
        G_2(m) = WorksheetFunction.ImProduct(WorksheetFunction.ImPower(TFactor_N2, k * m), f_2(m))
    Next m
    Result(k + 1, 2) = Val(WorksheetFunction.ImAbs(WorksheetFunction.ImSum(WorksheetFunction.ImSum(G_1), WorksheetFunction.ImProduct(WorksheetFunction.ImSum(G_2), WorksheetFunction.ImPower(TFactor_N1, k))))) / n
    Result(k + 1, 1) = k / (TimeLapse)
Next k
FFTFunction = Result
End Function
