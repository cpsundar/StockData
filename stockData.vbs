Sub stock_data()
Dim ws As Worksheet

Dim Great_per_change_inc_ticker As String
Dim Great_per_change_inc_value As Double

Dim Great_per_change_dec_ticker As String
Dim Great_per_change_dec_value As Double

Dim Greatest_volume_ticker_symbol As String
Dim Greatest_volume_total As Double



For Each ws In Worksheets
     'MsgBox ws.Name
    Dim LastRow As Double
    'Summarization
    Dim Sum_Row As Integer
    'Summarization headers per sheet
    Sum_Row = 1
    ws.Cells(Sum_Row, 9) = "Ticker"
    ws.Cells(Sum_Row, 10) = "Yearly Change"
    ws.Cells(Sum_Row, 11) = "Yearly Change Percentage"
    ws.Cells(Sum_Row, 12) = "Total Stock Volume"
 
      
    'MsgBox ws.Name
 
     
     Dim Ticker_First_day_open_price As Currency
     Dim Ticker_Last_day_close_price As Currency
     
     Dim Yearly_change As Double
     Dim Yearly_change_percentage As Double
     
     
     'Find the last Row of the current sheet
     LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'First price of the sheet which stored per year for the First stock symbol
      Ticker_First_day_open_price = ws.Cells(2, 3)
     
     'Loop through the Data Rows to check the Ticker symbos and calculate the Total
     For i = 2 To LastRow
            ' Assign the symbol
            ticker_symbol = ws.Cells(i, 1)
            ticker_symbol_next = ws.Cells(i + 1, 1)
            ticker_volume = ticker_volume + ws.Cells(i, 7)
            Ticker_Last_day_close_price = ws.Cells(i, 6)
            
            ' Check the next row for Ticket symbol change
            If (ticker_symbol <> ticker_symbol_next) Then
                   'If next is different symbol, write down Summarization of the current Ticker symbol
                   'Add next to last Summarization
                   Sum_Row = Sum_Row + 1
                   ws.Cells(Sum_Row, 9) = ticker_symbol
                   ws.Cells(Sum_Row, 12) = ticker_volume
                   'Reset the volume whenever the symbol changes
                   ticker_volume = 0
                   
                   ' Yearly change
                   Yearly_change = Ticker_Last_day_close_price - Ticker_First_day_open_price
            
                   
                   If Ticker_First_day_open_price = 0 Then
                        Yearly_change_percentage = 0
                   Else
                        Yearly_change_percentage = Yearly_change / Ticker_First_day_open_price
                   End If
                   
                   ws.Cells(Sum_Row, 10) = Yearly_change
                   ws.Cells(Sum_Row, 11) = Yearly_change_percentage
                   
                   
                   ws.Cells(Sum_Row, 10).NumberFormat = "0.000000000"
                   ws.Cells(Sum_Row, 11).NumberFormat = "0.0#%"
                   
                   
                   If Yearly_change > 0 Then
                     ws.Cells(Sum_Row, 10).Interior.Color = vbGreen
                   ElseIf Yearly_change < 0 Then
                     ws.Cells(Sum_Row, 10).Interior.Color = vbRed
                   End If
               
               ' Find the open price of the first day of the year
                Ticker_First_day_open_price = ws.Cells(i + 1, 3)
            End If
         
     Next i
 
   ' Hard
         ' initialize per sheet
    Great_per_change_inc_ticker = ws.Cells(2, 9)
    Great_per_change_inc_value = ws.Cells(2, 11)
    
    Great_per_change_dec_ticker = ws.Cells(2, 9)
    Great_per_change_dec_value = ws.Cells(2, 11)
    
    Greatest_volume_ticker_symbol = ws.Cells(2, 9)
    Greatest_volume_total = ws.Cells(2, 12)
    
 
     'Find the last Row of the current sheet
     LastRow_sum = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
     For sum_i = 3 To LastRow_sum
       ' To find Greatest percentage increase
       If ws.Cells(sum_i, 11) > Great_per_change_inc_value Then
        Great_per_change_inc_ticker = ws.Cells(sum_i, 9)
        Great_per_change_inc_value = ws.Cells(sum_i, 11)
       End If
       
           ' To find Greatest percentage decrease
       If ws.Cells(sum_i, 11) < Great_per_change_dec_value Then
        Great_per_change_dec_ticker = ws.Cells(sum_i, 9)
        Great_per_change_dec_value = ws.Cells(sum_i, 11)
       End If
       
         ' To find Greatest percentage decrease
       If ws.Cells(sum_i, 12) > Greatest_volume_total Then
        Greatest_volume_ticker_symbol = ws.Cells(sum_i, 9)
        Greatest_volume_total = ws.Cells(sum_i, 12)
       End If
       
       
     Next sum_i
     
     'Header
     ws.Cells(1, 15) = "Ticker"
     ws.Cells(1, 16) = "Value"
     
     'Greatest percentage increase
     ws.Cells(2, 14) = "Greatest % increase"
     ws.Cells(2, 15) = Great_per_change_inc_ticker
     ws.Cells(2, 16) = Great_per_change_inc_value
     ws.Cells(2, 16).NumberFormat = "0.00#%"
    
     'Greatest percentage decrease
     ws.Cells(3, 14) = "Greatest % decrease"
     ws.Cells(3, 15) = Great_per_change_dec_ticker
     ws.Cells(3, 16) = Great_per_change_dec_value
     ws.Cells(3, 16).NumberFormat = "0.00#%"
    
     'Greatest Total Volume
     ws.Cells(4, 14) = "Greatest Total Volume"
     ws.Cells(4, 15) = Greatest_volume_ticker_symbol
     ws.Cells(4, 16) = Greatest_volume_total
     ws.Cells(4, 16).NumberFormat = "##############"
    
    
Next ws



End Sub
