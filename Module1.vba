Attribute VB_Name = "Module1"

Sub Stocks()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

Dim TickerName As String
 
Dim VolumeTotal As Double
VolumeTotal = 0

Dim SummaryTable As Long
SummaryTable = 2

Dim LastRow As Long


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRow


    ws.Cells(1, 9).Value = "Ticker Name"
    ws.Cells(1, 10).Value = "Total Volume"
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    TickerName = ws.Cells(i, 1).Value
    VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
    
    ws.Range("I" & SummaryTable).Value = TickerName
    ws.Range("J" & SummaryTable).Value = VolumeTotal
    
    SummaryTable = SummaryTable + 1
    
    
    VolumeTotal = 0

Else

    VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub

