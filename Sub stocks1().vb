Sub stocks1()
' Define Variables
    Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Dim ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Dim Summary As Double
    Dim LastRow As Long
    Dim tickertracker As Double
    tickertracker = 0

'Title Columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'Summary Row
    Summary_Row = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Set initial open price

For i = 2 To LastRow
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'   Ticker Value
        ticker = Cells(i, 1).Value
'   YearChange Calculation
        YearChange = YearOpen - YearClose
        ' YearOpen = ws.Cells(i - tickertracker, 3).Value
        YearOpen = ws.Cells(i + 1, 3).Value
        YearClose = ws.Cells(i, 6).Value
        YearChange = YearOpen - YearClose
        If (YearChange = 0) Or (YearOpen = 0) Then
        ws.Range("K" & Summary_Row).Value = "null"
        Else
'   Percent Change Calculation
        PercentChange = ((YearClose - YearOpen) / YearOpen)
        ws.Columns("K").NumberFormat = "0.00%"
        End If
'   Volume
        Volume = ws.Cells(i, 7).Value
'   Return Values, Find Last Row
        ws.Cells(Summary_Row, 9).Value = ticker
        ws.Cells(Summary_Row, 10).Value = YearChange
        ws.Cells(Summary_Row, 11).Value = PercentChange
        ws.Cells(Summary_Row, 12).Value = Volume
        Summary_Row = Summary_Row + 1
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Volume = 0
    End If
'   Conditional Formatting
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
'   Finish loop
    Next i
    Next ws
End Sub