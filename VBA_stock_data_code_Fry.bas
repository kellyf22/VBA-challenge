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

'Start with an outer for loop to run the macro in each worksheet automatically
For Each Current In Worksheets

    'write the output headers to the row 1 cells and format Percent Change column as %
    Current.Cells(1, 9).Value = "Ticker"
    Current.Cells(1, 10).Value = "Yearly Change"
    Current.Cells(1, 11).Value = "Percent Change"
    Current.Range("K:K").NumberFormat = "0.00%"
    Current.Cells(1, 12).Value = "Total Volume"
    
    'Define 4 variables for our solutions
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Tot_Vol As Double
    
    'Define two more variables to keep track of values we'll need to find the solutions
    Dim year_open As Double
    Dim year_close As Double
    
    'lastrow1 is last row of provided data, tells us how many times we need to go through the loop
    lastrow1 = Current.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set initial conditions 
    Ticker = Current.Cells(2, 1).Value
    year_open = Current.Cells(2, 3).Value
    Tot_Vol = 0

    'Loop through all rows, calculate solutions, and write solutions in the same worksheet
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
            
            'record answers at end of list in columns to the right of provided data. lastrow2 is the last row in the columns of answers.
            lastrow2 = Current.Cells(Rows.Count, 9).End(xlUp).Row + 1
            Current.Cells(lastrow2, 9) = Ticker
            Current.Cells(lastrow2, 10) = Yearly_Change
            Current.Cells(lastrow2, 11) = Percent_Change
            Current.Cells(lastrow2, 12) = Tot_Vol
                           
            'increment or reset counters before beginning the next loop for a new stock.
            Ticker = Current.Cells(i + 1, 1).Value
            year_open = Current.Cells(i + 1, 3).Value
            Tot_Vol = 0
        Else
        
            'If the ticker symbol matched the previous one, adjust total volume only
            Tot_Vol = Tot_Vol + Current.Cells(i, 7).Value
        
        End If
      
    Next i
    
    'Bonus!
    'Greatest % increase is found with excel's Max function
    Current.Range("O2") = "Greatest % Increase"
    MyMax = Application.WorksheetFunction.Max(Current.Range("K:K"))

    'Return the row associated with the max percent in order to find the ticker using excel's Match function
    MaxCellRow = Application.WorksheetFunction.Match(MyMax, Current.Range("K:K"), 0)
    
    'Record answers and format as percent 
    Current.Range("Q2") = MyMax
    Current.Range("Q2").NumberFormat = "0.00%"
    Current.Range("P2") = Current.Cells(MaxCellRow, 9).Value
        
    'Greatest % decrease is found with excel's Min function
    Current.Range("O3") = "Greatest % Decrease"
    MyMin = Application.WorksheetFunction.Min(Current.Range("K:K"))
    MinCellRow = Application.WorksheetFunction.Match(MyMin, Current.Range("K:K"), 0)
    Current.Range("Q3") = MyMin
    Current.Range("Q3").NumberFormat = "0.00%"
    Current.Range("P3") = Current.Cells(MinCellRow, 9).Value
    
    'Greatest Total Volume is found with excel's Max function
    Current.Range("O4") = "Greatest Total Volume"
    MaxVol = Application.WorksheetFunction.Max(Current.Range("L:L"))
    MaxVolRow = Application.Match(MaxVol, Current.Range("L:L"), 0)
    Current.Range("Q4") = MaxVol
    Current.Range("P4") = Current.Cells(MaxVolRow, 9).Value
    
    'Put labels in appropriate cells
    Current.Range("P1") = "Ticker"
    Current.Range("Q1") = "Value"
    
Next

End Sub