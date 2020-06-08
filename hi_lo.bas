Attribute VB_Name = "Module1"
Sub Stocks_hi_lo():

    Dim volume_, i, j As Integer
    Dim open_, close_, change_, pchange_, max_pinc, max_pdec, max_volume, hi, lo, d_hi, d_lo As Double
    Dim mvt, mpit, mpdt As String
    
    For w = 1 To ThisWorkbook.Worksheets.Count
        
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
        lo = Range("E2")
        
        max_pinc = 0: max_pdec = 0: max_volume = 0
        
        While Cells(i - 1, 1) <> ""
        
            If Cells(i, 1) = Cells(i - 1, 1) Then
                volume_ = volume_ + Cells(i, 7)
                'adds to volume if ticker matches previous
                If Cells(i, 4) > hi Then
                    hi = Cells(i, 4)
                    d_hi = Cells(i, 2)
                End If
                If Cells(i, 5) < lo Then
                    lo = Cells(i, 5)
                    d_lo = Cells(i, 2)
                End If
                
            Else
                Cells(j, 12) = volume_
                'writes volume to cell
                If volume_ > max_volume Then
                    max_volume = volume_
                    mvt = Cells(i - 1, 1)
                End If
                volume_ = 0
                'resets volume
                close_ = Cells(i - 1, 6)
                change_ = close_ - open_
                If open_ > 0 Then
                    pchange_ = change_ / open_
                Else
                    pchange_ = 0
                End If
                'handled overflow error from stock PLNT, as all of its values were 0
                Cells(j, 10) = change_
                Cells(j, 11) = pchange_
                'calculates and writes total change and percent change to cells
                If pchange_ < max_pdec Then
                    max_pdec = pchange_
                    mpdt = Cells(i - 1, 1)
                ElseIf pchange_ > max_pinc Then
                    max_pinc = pchange_
                    mpit = Cells(i - 1, 1)
                End If
                
                Cells(j, 13) = hi
                Cells(j, 14) = d_hi
                Cells(j, 15) = lo
                Cells(j, 16) = d_lo
                
                hi = 0
                lo = Cells(i, 5)
                
                j = j + 1
                Cells(j, 9) = Cells(i, 1)
                'writes new ticker into next row
                open_ = Cells(i, 3)
                
            End If
                
            i = i + 1
            
        Wend
        
        Range("R2") = mpit
        Range("S2") = max_pinc
        Range("R3") = mpdt
        Range("S3") = max_pdec
        Range("R4") = mvt
        Range("S4") = max_volume
        'writing maximum values to cells
        Range("K:K, S2, S3").NumberFormat = "0.00%"
        
    Next w

End Sub

