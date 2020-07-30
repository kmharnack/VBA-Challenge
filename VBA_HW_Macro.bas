Attribute VB_Name = "Module1"
Sub main()
'Cycle through worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'set variables
Dim tickerColumn As Integer
tickerColumn = 1
Dim tickerName As String
tickerName = " "
Dim tickerVolume As Double
tickerVolume = 0
Dim openPrice As Double
openPrice = 0
Dim closePrice As Double
closePrice = 0
Dim yearChange As Double
yearChange = 0
Dim percentChange As Double
percentChange = 0
Dim summaryRow As Long
summaryRow = 2
Dim i As Long

 
'set headers
ws.Range("I1").Value = "TickerName"
ws.Range("J1").Value = "Year Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Stock Volume"
 
On Error Resume Next
openPrice = ws.Cells(2, 3).Value
For i = 2 To ws.Cells.SpecialCells(xlCellTypeLastCell).Row
 
'if ticker is different
If ws.Cells(i + 1, tickerColumn).Value <> ws.Cells(i, tickerColumn).Value Then

    'store the brand
        tickerName = ws.Cells(i, tickerColumn).Value
    'yearchange
        closePrice = ws.Cells(i, 6).Value
        yearChange = closePrice - openPrice
    'percentChange
        If openPrice <> 0 Then
                    percentChange = (yearChange / openPrice) 
                End If
    'store the total
        tickerVolume = tickerVolume + ws.Cells(i, 7).Value

    'write the ticker and volume and change
        ws.Range("I" & summaryRow).Value = tickerName
        ws.Range("J" & summaryRow).Value = yearChange
            'conditional formatting
                If (yearChange > 0) Then
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 4
                ElseIf (yearChange <= 0) Then
                    ws.Range("J" & summaryRow).Interior.ColorIndex = 3
                End If
     'write percent change and ticker volume
        ws.Range("K" & summaryRow).Value = percentChange
        ws.Range("K" & summaryRow).NumberFormat = "0.00%"
        ws.Range("L" & summaryRow).Value = tickerVolume
   
    'reset values
        summaryRow = summaryRow + 1
        yearChange = 0
        closePrice = 0
        openPrice = ws.Cells(i + 1, 3).Value
        tickerVolume = 0

'if ticker is same
Else
     tickerVolume = tickerVolume + ws.Cells(i, 7).Value
End If

Next i
Next ws
End Sub

              


