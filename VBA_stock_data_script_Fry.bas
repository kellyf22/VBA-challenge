Attribute VB_Name = "Module3"
Sub stocks3()

'for all stocks, output:
'1. ticker symbol
'2. yearly change in price = closing price on last trading day minus opening price on first trading day
'3. percent change in price = yearly change divided by opening price
'4. total stock volume = sum of all volumes for the year

'Bonus: Return stocks with
'1. Greatest % increase
'2. Greatest % decrease
'3. Greatest total volume

Dim Current As Worksheet

'Run the macro in each worksheet
For Each Current In Worksheets

    'write the output headers to the row 1 cells and format Percent Change column as %
    Current.Cells(1, 9).Value = "Ticker"
    Current.Cells(1, 10).Value = "Yearly Change"
    Current.Cells(1, 11).Value = "Percent Change"
    Current.Range("K:K").NumberFormat = "0.00%"
    Current.Cells(1, 12).Value = "Total Volume"
    
    'define some variables we'll need
    '4 required pieces of information
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Tot_Vol As Double
    
    'A few more variables to help us find the solutions
    Dim ticker_start_row As Double
    Dim year_open As Double
    Dim year_close As Double
    
    'lastrow1 is last row of provided data, tells us how many times we need to go through the loop
    lastrow1 = Current.Cells(Rows.Count, 1).End(xlUp).Row
    
    'initial conditions to begin the loop
    Ticker = Current.Cells(2, 1).Value
    ticker_start_row = 2
    year_open = Current.Cells(2, 3).Value
    Tot_Vol = 0
    
    For i = 2 To lastrow1
          
        'Find break points between stocks by comparing tickers in each row.
        'If the tickers are different, row i is the last row for one ticker, and i+1 is the first row of the next.
        If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
        
            year_close = Current.Cells(i, 6).Value
            
            'calculate answers
            Yearly_Change = year_close - year_open
            
            'When calculating Percent_Change, don't divide by zero!
            If year_open = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = Yearly_Change / year_open
            End If
            
            Tot_Vol = Tot_Vol + Current.Cells(i, 7).Value
            
            'record answers, lastrow2 is last row of output
            lastrow2 = Current.Cells(Rows.Count, 9).End(xlUp).Row + 1
            Current.Cells(lastrow2, 9) = Ticker
            Current.Cells(lastrow2, 10) = Yearly_Change
            Current.Cells(lastrow2, 11) = Percent_Change
            Current.Cells(lastrow2, 12) = Tot_Vol
                           
            'increment or reset counters
            Ticker = Current.Cells(i + 1, 1).Value
            year_open = Current.Cells(i + 1, 3).Value
            Tot_Vol = 0
        Else
        
            'If the ticker symbol matched the previous one, adjust total volume
            Tot_Vol = Tot_Vol + Current.Cells(i, 7).Value
        
        End If
      
    Next i
    
    'Bonus!
    'Greatest % increase
    Current.Range("O2") = "Greatest % Increase"
    MyMax = Application.WorksheetFunction.Max(Current.Range("K:K"))
    'Return the row associated with the max percent in order to find the ticker
    MaxCellRow = Application.WorksheetFunction.Match(MyMax, Current.Range("K:K"), 0)
    Current.Range("Q2") = MyMax
    Current.Range("Q2").NumberFormat = "0.00%"
    Current.Range("P2") = Current.Cells(MaxCellRow, 9).Value
        
    'Greatest % decrease
    Current.Range("O3") = "Greatest % Decrease"
    MyMin = Application.WorksheetFunction.Min(Current.Range("K:K"))
    MinCellRow = Application.WorksheetFunction.Match(MyMin, Current.Range("K:K"), 0)
    Current.Range("Q3") = MyMin
    Current.Range("Q3").NumberFormat = "0.00%"
    Current.Range("P3") = Current.Cells(MinCellRow, 9).Value
    
    'Greatest Total Volume
    Current.Range("O4") = "Greatest Total Volume"
    MaxVol = Application.WorksheetFunction.Max(Current.Range("L:L"))
    MaxVolRow = Application.Match(MaxVol, Current.Range("L:L"), 0)
    Current.Range("Q4") = MaxVol
    Current.Range("P4") = Current.Cells(MaxVolRow, 9).Value
    
    Current.Range("P1") = "Ticker"
    Current.Range("Q1") = "Value"
    
Next

End Sub



