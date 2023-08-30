Attribute VB_Name = "Module1"
Sub run_on_all_worksheets()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate

date_Format
ticker
yearchange
Greatestchange

Next ws

End Sub


Sub Greatestchange()

Dim great_increase As Double
Dim great_decrease As Double
Dim great_volume As LongLong

Range("P2").Value = "Greatest % increase"
Range("P3").Value = "Greatest % decrease"
Range("P4").Value = "Greatest Total Volume"
Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

great_increase = WorksheetFunction.Max(Range("M:M"))
great_decrease = WorksheetFunction.Min(Range("M:M"))
great_volume = WorksheetFunction.Max(Range("N:N"))

Range("R2").Value = great_increase
Range("R3").Value = great_decrease
Range("R4").Value = great_volume

Range("Q2").Value = Range("K" & (WorksheetFunction.Match(Range("R2"), Range("M:M"), 0)))
Range("Q3").Value = Range("K" & (WorksheetFunction.Match(Range("R3"), Range("M:M"), 0)))
Range("Q4").Value = Range("K" & (WorksheetFunction.Match(Range("R4"), Range("N:N"), 0)))


End Sub

Sub yearchange()

Dim yearchange As Double
Dim percentchange As Double
Dim volume As LongLong
Dim MaxDate As Date
Dim MinDate As Date
Dim openprice As Double
Dim closeprice As Double
Dim i As Long
Dim lastrow As Long


lastrow = Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To lastrow

MaxDate = WorksheetFunction.MaxIfs(Range("I:I"), Range("A:A"), Range("K" & i))
MinDate = WorksheetFunction.MinIfs(Range("I:I"), Range("A:A"), Range("K" & i))

'Range("P" & i).Value = MaxDate
'Range("Q" & i).Value = MinDate

openprice = WorksheetFunction.SumIfs(Range("C:C"), Range("I:I"), MinDate, Range("A:A"), Range("K" & i))
closeprice = WorksheetFunction.SumIfs(Range("F:F"), Range("I:I"), MaxDate, Range("A:A"), Range("K" & i))
volume = WorksheetFunction.SumIfs(Range("G:G"), Range("A:A"), Range("K" & i))

yearchange = closeprice - openprice
percentchange = yearchange / openprice

'Range("R" & i).Value = openprice
'Range("S" & i).Value = closeprice
Range("L" & i).Value = yearchange
Range("M" & i).Value = percentchange
Range("N" & i).Value = volume

Next i

End Sub

Sub ticker()

Dim i As Long
Dim lastrow As Long
Dim counter As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
counter = 2

For i = 2 To lastrow

If Range("A" & i).Value <> Range("A" & (i + 1)).Value Then

Range("K" & counter).Value = Range("A" & i)

counter = counter + 1

End If

Next i

End Sub

Sub date_Format()

Dim year As String
Dim month As String
Dim day As String
Dim date1 As Date
Dim lastrow As Long
Dim i As Long

Range("I1").Value = "Reconstructed date"
Range("K1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

year = Left(Range("B" & i).Value, 4)
day = Right(Range("B" & i).Value, 2)
month = Left(Right(Range("B" & i).Value, 4), 2)

date1 = month & "/" & day & "/" & year

Range("I" & i).Value = date1

Next i


End Sub

Sub Format()

End Sub


