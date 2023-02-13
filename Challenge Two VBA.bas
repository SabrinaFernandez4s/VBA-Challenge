Attribute VB_Name = "Module2"
Sub bb()
Dim ws As Worksheet
For Each ws In Worksheets


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim i As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'String because tickers are letters
Dim ticker As String
'for making sure that the tickers show up in one column
Dim tickerno As Integer
tickerno = 2
Dim TTLVol As Double
Dim OPN As Double
Dim CLSE As Double
Dim IndivVol As Double
Dim YrChng As Double
Dim PrcntChng As Double

OPN = ws.Range("C2").Value
IndivVol = 0
'start the for loop
For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value
    CLSE = ws.Cells(i, 6).Value
YrChng = CLSE - OPN
PrcntChng = YrChng / OPN
IndivVol = ws.Cells(i, 7).Value

'insertions
ws.Range("I" & tickerno).Value = ticker
ws.Range("J" & tickerno).Value = YrChng
ws.Range("K" & tickerno).Value = PrcntChng
ws.Range("L" & tickerno).Value = TTLVol + IndivVol

tickerno = tickerno + 1

Else
TTLVol = TTLVol + ws.Cells(i, 7).Value


End If
Next i
Next ws
Call Color
End Sub

Sub Color()
Dim ws As Worksheet
For Each ws In Worksheets

Dim i, LastRow As Double
LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For i = 2 To LastRow
    'if value is less than 0, make it red
    If ws.Cells(i, 10).Value < 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 3
      ws.Cells(i, 11).NumberFormat = "0.00%"
      
      'If value is more than or equal to zero, make it green
      ElseIf ws.Cells(i, 10).Value > 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
      ws.Cells(i, 11).NumberFormat = "0.00%"
      
    'If its any othervalue, make it no color
    Else: ws.Cells(i, 10).Interior.ColorIndex = xlNone
            ws.Cells(i, 11).NumberFormat = "0.00%"
End If
Next i
Next ws
Call Func
End Sub

Sub Func()
Dim ws As Worksheet
For Each ws In Worksheets

'define the variables
Dim i, LastRow As Integer

LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRow

'assign values

    If ws.Cells(i, 11).Value = WorksheetFunction.Max((ws.Range("L1:L" & LastRow).Value)) Then
    ws.Range("Q2").Value = ws.Cells(i, 11).Value
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"

End If
Next i
Next ws
End Sub
