
 'Running same macro on multiple sheets for years 2014 through 2016
   Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
    End Sub

Sub RunCode()

 'Introduce variables
    Dim ticker As String
    Dim open_price As Long
    Dim close_price As Long
    Dim high_stock As Long
    Dim low_stock As Long
    Dim daily_stock As Long
    Dim ws As Worksheet
    Dim i As Double
    Dim lastrow As Long

 'Assign Opening price, total_stock_value, and row_tot
    
    Dim opening_price_row As Double
    opening_price_row = 2
    
    Dim total_stock_vol As Double
    total_stock_vol = 0
    
    Dim row_tot As Integer
    row_tot = 2
    
 'Check the ticker value
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
         If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
         
            'Add unique ticker value
            Range("I" & row_tot).Value = Cells(i, 1).Value
                    
            'Calculate yearly change
            
            Range("J" & row_tot).Value = Cells(i, 6).Value - Cells(opening_price_row, 3).Value
            
            'Calculate percentage change
            
            If Cells(opening_price_row, 3).Value <> 0 Then
            
                Range("K" & row_tot).Value = (Cells(i, 6).Value - Cells(opening_price_row, 3)) / (Cells(opening_price_row, 3))
            Else
                Range("K" & row_tot).Value = Cells(i, 6).Value
            End If
            
 'Iterate row for opening_price
 
            opening_price_row = i + 1 
                
 'Calculate total stock
            
            Range("L" & row_tot).Value = total_stock_vol + Cells(i, 7).Value
                
            row_tot = row_tot + 1
            
            'Reset the totalstock volume
                
            total_stock_vol = 0
                
        Else
            total_stock_vol = total_stock_vol + Cells(i, 7).Value
                    
        End If
                    
                
    Next i

  'Conditional Formatting

   lastrow = Cells(Rows.Count, 10).End(xlUp).Row

            For i = 2 To lastrow

                If Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 3

                ElseIf Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4

                ElseIf Cells(i, 10).Value = 0 Then
                    Cells(i, 10).Interior.ColorIndex = xlNone

                End If
            Next i
                

           
  'Find Growth, Loss in stock (based on percentage change column)and Greatest Increase in total stock volume

    
    For Each ws In Worksheets

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        lastrow = ws.Range("K" & Rows.Count).End(xlUp).Row
        
 'Find the maximum and the minimum number
        max_num = 0
        min_num = 0
        vol_num = 0

        For i = 2 To lastrow
            If ws.Range("K" & i).Value > max_num Then
                max_num = ws.Range("K" & i).Value
                max_ticker_row = i
                
            ElseIf ws.Range("K" & i).Value < min_num Then
                min_num = ws.Range("K" & i).Value
                min_ticker_row = i
                
            ElseIf ws.Range("L" & i).Value > vol_num Then
                vol_num = ws.Range("L" & i).Value
                vol_ticker_row = i
            End If
            
        Next i
        
        ws.Range("P2").Value = ws.Range("I" & max_ticker_row).Value
        ws.Range("Q2").Value = max_num
        
        ws.Range("P3").Value = ws.Range("I" & min_ticker_row).Value
        ws.Range("Q3").Value = min_num
        
        ws.Range("P4").Value = ws.Range("I" & vol_ticker_row).Value
        ws.Range("Q4").Value = vol_num
        
 'Formatting
        ws.Range("Q2:Q3" & new_lastrow).NumberFormat = "0.00%"
        
    Next ws
    
End Sub







