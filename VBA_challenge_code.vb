'run this subroutine to run code on all worksheets
Public Sub Main_Run()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate
        Call stock_changes(ws)
        
    Next ws
    MsgBox ("Done!")

End Sub

'this subroutine is the looping code that does the work
Sub stock_changes(ws As Worksheet)

    'define all variables
    Dim ticker As String
    Dim year_start As Double
    Dim year_end As Double
    Dim volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim summary_table_row As Integer
    Dim corresponding_ticker As String
    Dim max_total_volume As Double
    
    
    'determine the last row with data in the worksheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'define initial variable values
    summary_table_row = 2
    year_start = Cells(2, 3).Value
    volume = 0
    
    
    '===========================
    'INITIAL DATA GATHERING LOOP
    '===========================
    
    
    'loop through all the rows of the sheet
    For i = 2 To lastrow
        
        'check if the ticker in the next row is different from the current row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'ticker has changed, get previous ticker info from data and add to summary table
            ticker = Cells(i, 1).Value
            volume = volume + Cells(i, 7).Value
            year_end = Cells(i, 6).Value
            yearly_change = year_end - year_start
            percent_change = yearly_change / year_start
            
            'input values into table
            Cells(summary_table_row, 9).Value = ticker
            Cells(summary_table_row, 10).Value = yearly_change
            Cells(summary_table_row, 11).Value = percent_change
            Cells(summary_table_row, 12).Value = volume
            
            'format percent_change column as percentage
            Cells(summary_table_row, 11).NumberFormat = "0.00%"
            
            'round yearly_change column to 2 decimal places
            Cells(summary_table_row, 10).NumberFormat = "0.00"
            
            'increase the summary table row by 1
            summary_table_row = summary_table_row + 1
            
            'reset the volume to 0 for the next loop
            volume = 0
            
            'reset year_start to value from next ticker
            year_start = Cells(i + 1, 3)
            
        'ticker hasn't changed
        Else
            
            'add volume for current row to volume total
            volume = volume + Cells(i, 7).Value
            
            
        End If
        
    'next loop
    Next i
        
        
    '========================
    'GREATEST % INCREASE LOOP
    '========================
    
    
    'initialize variables
    max_percent_increase = 0
    ticker = Cells(2, 9).Value
    corresponding_ticker = ticker
    
    'loops through all the rows of the sheet
    For i = 2 To lastrow
        
        'check if the ticker in the next row is different from the current row
        If Cells(i + 1, 9).Value <> Cells(i, 9).Value Then
            
            'ticker has changed, check if new percent change is greater than old percent change
            If max_percent_increase < percent_change Then
                max_percent_increase = percent_change
                'update ticker when new max is found
                corresponding_ticker = ticker
            End If
            
            'update ticker and reset for new ticker
            ticker = Cells(i, 9).Value
            percent_change = Cells(i, 11).Value
            
        End If
            
        'update percent_change for each new row
        percent_change = Cells(i, 11).Value
    
    Next i
        
    'print the greatest percent increase and its ticker to cells
    Cells(1, 16).Value = "Ticker"
    Cells(2, 16).Value = corresponding_ticker
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 17).Value = max_percent_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(1, 17).Value = "Value"
    
    '========================
    'GREATEST % DECREASE LOOP
    '========================
    
    'initialize variables
    max_percent_decrease = 0
    ticker = Cells(2, 9).Value
    corresponding_ticker = ticker
    
    'loops through all the rows of the sheet
    For i = 2 To lastrow
        
        'check if the ticker in the next row is different from the current row
        If Cells(i + 1, 9).Value <> Cells(i, 9).Value Then
            
            'ticker has changed, check if new percent change is less than old percent change
            If max_percent_decrease > percent_change Then
                max_percent_decrease = percent_change
                'update ticker when new max is found
                corresponding_ticker = ticker
            End If
            
            'update ticker and reset for new ticker
            ticker = Cells(i, 9).Value
            percent_change = Cells(i, 11).Value
            
        End If
            
        'update percent_change for each new row
        percent_change = Cells(i, 11).Value
    
    Next i
    
    'print the greateast percent decrease and its ticker into cells
    Cells(3, 16).Value = corresponding_ticker
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 17).Value = max_percent_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    
    
    '==========================
    'GREATEST TOTAL VOLUME LOOP
    '==========================
    
    'initialize variables
    max_total_volume = 0
    ticker = Cells(2, 9).Value
    corresponding_ticker = ticker
    
    'loops through all the rows of the sheet
    For i = 2 To lastrow
    
        'check if the ticker in the next row is different from the current row
        If Cells(i + 1, 9).Value <> Cells(i, 9).Value Then
            
            'ticker has changed, check if new total stock volume is greater than old stock volume
            If max_total_volume < volume Then
                max_total_volume = volume
                'update ticker when new max is found
                corresponding_ticker = ticker
            End If
            
            'update ticker and reset for new ticker
            ticker = Cells(i, 9).Value
            volume = Cells(i, 12).Value
            
        
        End If
        
        'update volume for each new row
        volume = Cells(i, 12).Value
            
    Next i
    
    'print the greatest total volume and its ticker into cells
    Cells(4, 16).Value = corresponding_ticker
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 17).Value = max_total_volume
    
    
    '=============================
    'AFTER DATA PROCESSING IS DONE
    '=============================
           
    'add headings to summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'loop through yearly change column and format with red for negative change and green for positive change
    For i = 2 To lastrow
        
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = vbGreen
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.Color = vbRed
        End If
        
    Next i

    'loop through percent change column and format with red for negative change and green for positive change
    For i = 2 To lastrow
        
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.Color = vbGreen
        ElseIf Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.Color = vbRed
        End If
        
    Next i

    
    'autofit columns for readability
    Range("I1:L3005").Columns.AutoFit
    Range("O1:O4").Columns.AutoFit
    Range("P1:Q3").Columns.AutoFit
    
End Sub

