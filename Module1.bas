Attribute VB_Name = "Module1"
Sub testscript()
   MsgBox Application.Evaluate("=CELL(""address"",INDEX(A4:H4,MATCH(MIN(A4:H4),A4:H4,0)))")
End Sub


Sub StockMarketVBA_Final()

Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Stock Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



Dim symbol As String
Dim symbolNumber As Long
Dim FinalRow As Long

symbol = " "
symbolNumber = 1
FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim start, finish, totalPrice As Double
start = 0
finish = 0
totalPercent = 0
totalPrice = 0


Dim bestPerformer As Double
Dim worstPerformer As Double
Dim r As Range

Dim output As Integer


output = 2

    For i = 2 To FinalRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        symbolNumber = symbolNumber + 1
        symbol = ws.Cells(i, 1).Value
        
        ws.Cells(symbolNumber, 9).Value = symbol
        start = start + ws.Cells(i, 3).Value
        finish = finish + ws.Cells(i, 6).Value
        totalPrice = start - finish
        total_stock_vol = total_stock_vol + ws.Cells(i, 7)
        
        If (Abs(finish)) > 0 Then
        totalPercent = ((totalPrice) / (Abs(finish))) * 100
        End If
        ws.Range("L" & output).Value = total_stock_vol
        ws.Range("J" & output).Value = totalPrice
        ws.Range("K" & output).Value = totalPercent
        ws.Range("K" & output).NumberFormat = "0.00%"

        If ws.Range("J" & output).Value < 0 Then
        ws.Range("J" & output).Interior.ColorIndex = 3 ' Red
            
        Else:
        ws.Range("J" & output).Interior.ColorIndex = 4 ' Green
        End If
        output = output + 1

        start = 0
        finish = 0
        totalPrice = 0
        totalPercent = 0
        total_stock_vol = 0
        
        Else:
        start = start + ws.Cells(i, 3).Value
        finish = finish + ws.Cells(i, 6).Value
        totalPrice = start - finish
        
        If (Abs(finish) > 0) Then
        totalPercent = ((totalPrice) / (Abs(finish))) * 100
        End If
        
        total_stock_vol = total_stock_vol + ws.Cells(i, 7)
        
    
        End If
        
        
        
    Next i
    

    
Next ws




End Sub

Function GetAddr(rng As Range) As Range


    Dim dMin As Double
    Dim lIndex As Long
    Dim sAddress As String

    Application.Volatile
    With Application.WorksheetFunction
        dMin = .Min(rng)
        lIndex = .Match(dMin, rng, 0)
    End With
  Set GetAddr = rng.Cells(lIndex).Address
End Function
