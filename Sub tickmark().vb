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
    Dim Summary As Integer
'Title Columns
    ws.Cells(1,9).Value = "Ticker"
    ws.Cells(1,10).Value="Yearly Change"
    ws.Cells(1,11).Value = "Percent Change"
    ws.Cells(1,12).Value ="Total Stock Volume"
'Summary Row
    Summary_Row = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i=2 to LastRow 
    If (YearChange = 0) Or (YearOpen = 0) Then
    ws.Range("K" & Summary_Row).value = "null"
    End If
    Next i 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'   Ticker Value
        ticker=Cells(i,1).value
'   YearChange Calculation
        YearOpen = ws.Cells(i,3).value
        YearClose = ws.Cells(i,6).value
        YearChange = YearOpen - YearClose
'   Percent Change Calculation
        PercentChange = (YearClose / YearOpen) 
        Columns("K").NumberFormat = "0.00%"
'   Volume
        Volume = ws.Cells(i, 7).value
'   Return Values, Find Last Row
        ws.Cells(Summary_Row,9).Value = ticker
        ws.cells(Summary_Row,10).Value = YearChange
        ws.Cells(Summary_Row,11).Value = PercentChange
        ws.Cells(Summary_Row,12).Value = Volume
        Summary_Row = Summary_Row + 1
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Volume = 0 
    End If
'   Finish loop
    Next i 
    Next ws
End Sub