Sub StocksBro()

    'Set variables
    Dim tickerName As String 'Hold ticker symbol
    Dim TotalVol As Double 'Hold total volume
    Dim lastRow As Double 'Count number of rows
    Dim yrChange As Double 'to calculate yearly change
    Dim change As Single 'to calculate % change
    Dim rowStart As Double 'Make the row iteration start at 2
    
    'Set New headers for Analysis
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    
    'Set titles for Bonus
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    
    'Set variables for Next loop - BONUS
    Dim greatTkr As String
    Dim greatPerInc As Double
    Dim greatPerDec As Double
    Dim greatVol As Double
        
    'Give variables initial value to hold
    TotalVol = 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    lastRow2 = Cells(Rows.Count, "K").End(xlUp).Row
    rowStart = 2
    
    'Keep track of the location of each ticker
    Dim SumTblRow As Integer
    SumTblRow = 2
    
    
    'Loop through all tickers
    For i = 2 To lastRow
            
        'iterate through each ticker and take note of those that do not equal each other
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            '-------------set calculations------------
            
            'Set the ticker name
            tickerName = Cells(i, 1).Value
            
            'Sum up total volume for that the ticker selected
            TotalVol = TotalVol + Cells(i, 7).Value
            
            'calculate yearly change
            yrChange = (Cells(i, 6) - Cells(rowStart, 3))
            
            
            'calculate change %
            change = Round(((Cells(i, 6) - Cells(rowStart, 3)) / Cells(rowStart, 3) * 100), 2)
            
            
            
            '-------------Print calculations------------
            'Print ticker Name in Column "J"
            Range("J" & SumTblRow).Value = tickerName
            Range("K" & SumTblRow).Value = yrChange
            Range("M" & SumTblRow).Value = TotalVol
            Range("L" & SumTblRow).Value = change & "%"
            
            
            'Make sure the row numbers for the print increments by 1!! <this is key
            SumTblRow = SumTblRow + 1
            
            
            TotalVol = 0
        
       Else
       
            'Add to total volume when the tickers equal each other - for a sumif effect
            TotalVol = TotalVol + Cells(i, 7).Value
       
       End If
       
     Next i
           
    'Conditional formatting for Yearly Change column
    For j = 2 To lastRow2
        If Cells(j, 11).Value > 0 Then
            Cells(j, 11).Interior.ColorIndex = 4
        Else
            Cells(j, 11).Interior.ColorIndex = 3
            
        End If
        
     Next j
           
     '---------------------------------------BONUS-----------------------
     'Get Greatest volumn/Per Increase/Per Decrease calculation
     greatVol = Application.WorksheetFunction.Max(Range("M:M"))
     greatPerInc = WorksheetFunction.Max(Range("L2:L" & lastRow))
     greatPerDec = WorksheetFunction.Min(Range("L2:L" & lastRow))
     
     'print Greatest Volume
     Range("R4").Value = greatVol
     Range("R3").Value = "%" & greatPerDec * 100
     Range("R2").Value = "%" & greatPerInc * 100
         
     'Tickers for greatest value
     For k = 2 To lastRow
        If Cells(k, 12).Value = greatPerInc Then
            Range("Q2").Value = Cells(k, 10)
        ElseIf Cells(k, 12).Value = greatPerDec Then
            Range("Q3").Value = Cells(k, 10)
        ElseIf Cells(k, 13).Value = greatVol Then
            Range("Q4").Value = Cells(k, 10)
        End If

     Next k
    
     
         
End Sub

