Attribute VB_Name = "Module1"
Sub CalculateQuarterlyChanges()

    Dim ws As Worksheet
    Dim ws_last_row As Double
    Dim ws_ticker As String
    Dim ws_open_price As Double
    Dim ws_close_price As Double
    Dim ws_quarterly_change As Double
    Dim ws_percentage_change As Double
    Dim ws_total_stock_volume As Double
    Dim ws_loop_counter As Integer
    Dim row As Double
    Dim column As Integer
    
    
     For Each ws In ActiveWorkbook.Worksheets
                    ws.Activate
    
    ' Find the last row with data in column A
    ws_last_row = Cells(Rows.Count, 1).End(xlUp).row
     
    
    ' Add headings to the newly created columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest Increase"
    Cells(3, 15).Value = "Greatest Percentage Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

 
        
    ws_total_stock_volume = 0
    row = 2
    column = 1
    
     ws_open_price = Cells(2, column + 2).Value
        
        
     

  
    'start the loop to determine the above values
      For i = 2 To ws_last_row
    
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
     
            ws_ticker = Cells(i, column).Value
            Cells(row, column + 8).Value = ws_ticker
    
            ws_close_price = Cells(i, column + 5).Value
                
             
            ws_quarterly_change = ws_close_price - ws_open_price
            Cells(row, column + 9).Value = ws_quarterly_change
               
                
            ws_percentage_change = ws_quarterly_change / ws_open_price
            Cells(row, column + 10).Value = ws_percentage_change
            Cells(row, column + 10).NumberFormat = "0.00%"

            ws_total_stock_volume = ws_total_stock_volume + Cells(i, column + 6).Value
            Cells(row, column + 11).Value = ws_total_stock_volume
        
    '        Reset and Count
            row = row + 1
            ws_open_price = Cells(i + 1, column + 2)
            ws_total_stock_volume = 0
            ws_perentage_change = 0
            ws_quarterly_change = 0
   
        Else
            ws_total_stock_volume = ws_total_stock_volume + Cells(i, column + 6).Value
        
        End If
      
      Next i

        
              
        
        ' find the last row of ticker column
        ws_quarterly_change_last_row = Cells(Rows.Count, 9).End(xlUp).row
        
        ' set the Cell Colors
        For P = 2 To ws_quarterly_change_last_row
            If (Cells(P, 10).Value > 0 Or Cells(P, 10).Value = 0) Then
                Cells(P, 10).Interior.ColorIndex = 10
            ElseIf Cells(P, 10).Value < 0 Then
                Cells(P, 10).Interior.ColorIndex = 3
            End If
        Next P
        
        ' find the highest value of each ticker
        For k = 2 To ws_quarterly_change_last_row
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & ws_quarterly_change_last_row)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & ws_quarterly_change_last_row)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ws_quarterly_change_last_row)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k

        
         
    Next ws

End Sub




