VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks():

    Dim volume_, i, j As Integer
    Dim open_, close_, change_, pchange_, max_pinc, max_pdec, max_volume As Double
    Dim mvt, mpit, mpdt As String
    'tickers for max volume, % increase, and % decrease respectively
    
    For w = 1 To ThisWorkbook.Worksheets.Count
        ThisWorkbook.Worksheets(w).Activate
    'loops through worksheets
    
        i = 3
        'i is the row index for our loop
        j = 2
        'j is the row index for the output
        open_ = Range("C2"): volume_ = Range("G2")
        'gives initial values to open, volume
        Cells(j, 9) = Range("A2")
        'places the first ticker
        
        max_volume = 0: max_pdec = 0: max_pinc = 0
        'resets max values between worksheets
        
        While Cells(i - 1, 1) <> ""
        'loop runs until it encounters a blank cell in the ticker column
            If Cells(i, 1) = Cells(i - 1, 1) Then
                volume_ = volume_ + Cells(i, 7)
                'adds to volume if ticker matches previous
                
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
                Cells(j, 10) = change_: Cells(j, 11) = pchange_
                'calculates and writes total change and percent change to cells
                If pchange_ < max_pdec Then
                    max_pdec = pchange_
                    mpdt = Cells(i - 1, 1)
                ElseIf pchange_ > max_pinc Then
                    max_pinc = pchange_
                    mpit = Cells(i - 1, 1)
                End If
                j = j + 1
                Cells(j, 9) = Cells(i, 1)
                'writes new ticker into next row
                open_ = Cells(i, 3)
                
            End If
                
            i = i + 1
            
        Wend
        
        Range("P2") = mpit: Range("Q2") = max_pinc
        Range("P3") = mpdt: Range("Q3") = max_pdec
        Range("P4") = mvt: Range("Q4") = max_volume
        'writing maximum values to cells
        Range("K:K, Q2, Q3").NumberFormat = "0.00%"
        
    Next w

End Sub

