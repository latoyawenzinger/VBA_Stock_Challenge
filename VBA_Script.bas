Attribute VB_Name = "Module1"
Sub StockLoop()

Application.ScreenUpdating = False
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate


'creating headers

ws.range("I1").Value = "Ticker"
ws.range("J1").Value = "Yearly Change"
ws.range("K1").Value = "Percent Change"
ws.range("L1").Value = "Total Stock Volume"

'assign and set variables

Dim ticker As String
Dim openprice, closedprice, percentchange, vol, maxincrease, mindecrease, maxvol As Double
vol = 0
Symbol = 2
maxincrease = 0
mindecrease = 0
maxvolume = 0


'loops through all ticker symbols
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

'set intitial openprice row
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

openprice = ws.Cells(i, 3).Value

End If


    'loops through ticker symbols and closed/open prices noting changes in ticker symbol
    If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then


    'sorting ticker symbol
    ticker = ws.Cells(i, 1).Value
    ws.range("I" & Symbol).Value = ticker

    'calculating year change
    closedprice = ws.Cells(i, 6).Value
    yearchange = closedprice - openprice
    ws.range("J" & Symbol).Value = yearchange
    
        
    
    'calculating percent change
    percentchange = (yearchange) / openprice
    ws.range("K" & Symbol).Value = percentchange
    ws.range("K" & Symbol).NumberFormat = "0.00%"
    
    

    'calculating volume
    vol = vol + ws.Cells(i, 7).Value
    ws.range("L" & Symbol) = vol


    Symbol = Symbol + 1
    vol = 0

' if next row immediatley below is the same symbol, then keep adding
    Else
    vol = vol + ws.Cells(i, 7).Value



End If

Next i

        'color coding negative vs postive nnumbers
        lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For j = 2 To lastrow
        'if negative, then red
        If ws.Cells(j, 10) <= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    
        'if positive, then green
        ElseIf ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
        
        End If
        
        
Next j

'creating column/row headers for next summary table
ws.range("O2") = "Greatest % Increase"
ws.range("O3") = "Greatest % Decrease"
ws.range("O4") = "Greatest Total Volume"
ws.range("P1") = "Ticker"
ws.range("Q1") = "Value"


'finding max and min% value from Column K


ws.Cells(2, 17).Value = WorksheetFunction.Max(range("K:K"))
maxincrease = ws.Cells(2, 17).Value
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 17).Value = WorksheetFunction.min(range("K:K"))
mindecrease = ws.Cells(3, 17).Value
ws.Cells(3, 17).NumberFormat = "0.00%"


lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
For x = 2 To lastrow

'define name for max/min % value

    If ws.Cells(x, 11).Value = maxincrease Then
        maxincrease = ws.Cells(x, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(x, 9).Value

    End If
    
Next x

For y = 2 To lastrow
 
    If ws.Cells(y, 11).Value = mindecrease Then
        mindecrease = ws.Cells(y, 11)
        ws.Cells(3, 16).Value = ws.Cells(y, 9).Value



     
    End If
    
Next y

lastrow = ws.Cells(Rows.Count, 12).End(xlUp).Row
For Z = 2 To lastrow

'finding max and min% value from Column L
ws.Cells(4, 17).Value = WorksheetFunction.Max(range("L:L"))
maxvol = ws.Cells(4, 17).Value

'finding max total volume
    If maxvol = ws.Cells(Z, 12).Value Then
        maxvol = ws.Cells(Z, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(Z, 9).Value
        
        

    End If

Next Z

'loop to next worksheet
Next ws

End Sub
