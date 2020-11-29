Attribute VB_Name = "Module1"
Sub Stocks():

    Dim volume_, i, j As Integer
    Dim open_, close_, change_, percent_change, max_percent_inc, max_percent_dec, max_volume, high, low, date_of_high, date_of_low As Double
    Dim max_volume_ticker, max_percent_inc_ticker, max_percent_dec_ticker As String
    
    For w = 1 To ThisWorkbook.Worksheets.Count
    'loops through worksheets, w is the worksheet index
        
        ThisWorkbook.Worksheets(w).Activate
        i = 3
        'i is the row index for our loop
        j = 2
        'j is the row index for the output
        open_ = Range("C2")
        volume_ = Range("G2")
        'gives initial values to open, volume
        Cells(j, 9) = Range("A2")
        'places the first ticker
        low = Range("E2")
        
    max_percent_inc = 0: max_percent_dec = 0: max_volume = 0
    'resets max values between worksheets
        
        While Cells(i - 1, 1) <> ""
        'loop runs until it encounters a blank cell in the ticker column
        
            If Cells(i, 1) = Cells(i - 1, 1) Then
                volume_ = volume_ + Cells(i, 7)
                'adds to volume if ticker matches previous
                If Cells(i, 4) > high Then
                    high = Cells(i, 4)
                    date_of_high = Cells(i, 2)
                End If
                'determining if current value is the high or low for particular ticker
                If Cells(i, 5) < low Then
                    low = Cells(i, 5)
                    date_of_low = Cells(i, 2)
                End If
                
            Else
                Cells(j, 12) = volume_
                'writes volume to cell
                If volume_ > max_volume Then
                    max_volume = volume_
                    max_volume_ticker = Cells(i - 1, 1)
                End If
                volume_ = 0
                'resets volume
                close_ = Cells(i - 1, 6)
                change_ = close_ - open_
                If open_ > 0 Then
                percent_change = change_ / open_
                Else
                percent_change = 0
                End If
                'handled overflow error from stock PLNT, as all of its values were 0
            
                Cells(j, 10) = change_
                Cells(j, 11) = percent_change
                'calculates and writes total change and percent change to cells
            
            If percent_change < max_percent_dec Then
                max_percent_dec = percent_change
                    max_percent_dec_ticker = Cells(i - 1, 1)
                'checks if percent changes are universal minimum or maximum
            ElseIf percent_change > max_percent_inc Then
                max_percent_inc = percent_change
                    max_percent_inc_ticker = Cells(i - 1, 1)
                End If
                
            'writing outputs for particular ticker
                Cells(j, 13) = high
                Cells(j, 14) = date_of_high
                Cells(j, 15) = low
                Cells(j, 16) = date_of_low
                
            'resets high and low
                high = 0
                low = Cells(i, 5)
                
                j = j + 1
                Cells(j, 9) = Cells(i, 1)
                'writes new ticker into next row
                open_ = Cells(i, 3)
                
            End If
                
            i = i + 1
            
        Wend
        
        'displaying universal maximum/minimum values
        Range("R2") = max_percent_inc_ticker
        Range("S2") = max_percent_inc
        Range("R3") = max_percent_dec_ticker
        Range("S3") = max_percent_dec
        Range("R4") = max_volume_ticker
        Range("S4") = max_volume
        
        'formatting appropriate columns as percentages
        Range("K:K, S2, S3").NumberFormat = "0.00%"
        
    Next w

End Sub

