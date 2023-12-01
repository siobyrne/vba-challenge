Sub main()
    ' iterate over worksheets and run createAnalysis on them
    ' I used this site ( https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html )
    ' as a code refernce for this section!
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call createAnalysis
    Next
    Application.ScreenUpdating = True
End Sub

Sub createAnalysis()
    ' assign names to stock info cells and set width as needed
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    Columns("J").ColumnWidth = 13
    [K1] = "Percentage Change"
    Columns("K").ColumnWidth = 14
    [L1] = "Total Stock Volume"
    Columns("L").ColumnWidth = 17
    [O2] = "Greatest % Increase"
    [O3] = "Greatest % Decrease"
    [O4] = "Greatest Total Volume"
    Columns("O").ColumnWidth = 20
    [P1] = "Ticker"
    [Q1] = "Value"
    Columns("Q").ColumnWidth = 17
    
    ' variables for ticker and sheet length
    Dim ticker As String: ticker = Cells(2, 1).Value
    Dim sheet_len As Double: sheet_len = Cells(Rows.Count, "A").End(xlUp).Row
    ' create rows
    Dim row_counter As Long: row_counter = 2
    Dim opening_price As Double: opening_price = Cells(2, 3).Value
    Dim closing_price As Double:
    Dim opening_vol As Range: Set opening_vol = Cells(2, 7)
    Dim closing_vol As Range

    ' loop to populate totals for ticker/yearly change/percent change/total stock volume columns
    Cells(row_counter, 9).Value = ticker
    For i = 2 To (sheet_len + 1)
        Dim cell As String: cell = Cells(i, 1).Value
        
        ' compares value in ticker to value in new ticker column to see if equal 
        ' and set values for stock info cells in createAnalysis()
        If Not cell = ticker Then
            closing_price = Cells(i - 1, 6).Value
            Set closing_vol = Cells(i - 1, 7)
            Cells(row_counter, 10).Value = closing_price - opening_price
            Cells(row_counter, 11).Value = (closing_price - opening_price) / opening_price
            Cells(row_counter, 12).Value = WorksheetFunction.Sum(Range(opening_vol.Address(), closing_vol.Address()))
            
            ' create new opening_price and opening_vol
            opening_price = Cells(i, 3).Value
            Set opening_vol = Cells(i, 7)
            
            ' increase row_counter and set new ticker variable
            row_counter = row_counter + 1
            ticker = cell
            Cells(row_counter, 9).Value = ticker
        End If
    Next i

 ' formatting 
    Dim table_len As Double: table_len = Cells(Rows.Count, 9).End(xlUp).Row
    Dim col_range As Range

    ' format yearly change 
    Set col_range = Range("J2:J" & table_len)
    col_range.NumberFormat = "0.00"

    ' set yearly change column colors to red and green for decrease/increase
    For i = 2 To table_len
        If Cells(i, 10).Value >= 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    ' format percentage change
    Set col_range = Range("K2:K" & table_len)
    col_range.NumberFormat = "0.00%"

 ' variables for table of greatest increase/decrease/total volume
    Dim increase As Double: increase = Cells(2, 11).Value
    Dim decrease As Double: decrease = Cells(2, 11).Value
    Dim volume As Double: volume = Cells(2, 12).Value

    ' populate ticker, increase/decrease/volume columns
    Cells(2, 16).Value = Cells(2, 9).Value
    Cells(3, 16).Value = Cells(2, 9).Value
    Cells(4, 16).Value = Cells(2, 9).Value
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = volume

    ' add percentage format for greatest increase/decrease column
    Set col_range = Range("Q2:Q3")
    col_range.NumberFormat = "0.00%"
    
    ' iterate down rows and update values in stock analysis box
    For i = 2 To table_len
        If Cells(i, 11) > increase Then
            increase = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = increase
        ElseIf Cells(i, 11).Value < decrease Then
            decrease = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = decrease
        ElseIf Cells(i, 12).Value > volume Then
            volume = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = volume
        End If
    Next i


End 
