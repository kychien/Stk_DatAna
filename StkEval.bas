Attribute VB_Name = "XLS_Stock_Eval"

Sub StkEval():
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to cycle through several excel sheets to gather stock
    '   volume over the years.
    '
    '   2018 08 13 - Added in implementation to check for data across multiple worksheets.
    '       Added in checks for IPO's. Fixed an error where the last ticker was not being accounted
    '       for.
    '   2018 08 10 - Basic implementation for 1 sheet of data to analyze total volume, overall
    '       change for the year, and percentage change for a given ticker symbol.  Also tracks
    '       greatest percent increase and decrease as well as greatest total volume
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim tSym() As String                        'Array to hold ticker symbols for aggregating data over years
    Dim tVol() As Double                        'Array to track corresponding volume
    Dim tOpn() As Double                        'Array to hold year open value
    Dim tCls() As Double                        'Array to hold close value at year end
    Dim tX, vX, pX, dY, oX, cX, size, tPos, IPOs As Integer
    Dim gpd, gpi, gtv As Integer
    Dim tCur, lAdr, rAdr As String
    Dim gpdT, gpiT, gtvT As String
    Dim notPublic As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Counter for current number of worksheets just in case future implementations create a new summary
    '   worksheet that shouldn't be evaluated
    Dim sheets As Integer
    sheets = Worksheets.Count
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '''''''''''''''''''''''''''''''''''''''''''''Loop for cycling through worksheets
    For n = 1 To sheets
        Worksheets(n).Activate
    
        '''''''''''''''''''''''''''''''''''''''''''''Initialize excel coordiantes...
        tX = WorksheetFunction.Match("<ticker>", Range("1:1"), 0)   ' ticker column via match
        dY = 2                                                      ' ticker row
        vX = WorksheetFunction.Match("<vol>", Range("1:1"), 0)      ' volume column via match
        pX = 1                                                      ' printout column
        oX = WorksheetFunction.Match("<open>", Range("1:1"), 0)     ' open column via match
        cX = WorksheetFunction.Match("<close>", Range("1:1"), 0)    ' close column via match
        size = 1                                                    ' number of unique ticker symbols
        tPos = 0                                                    ' array position tracker
        tCur = Cells(dY, tX)                                        ' current ticker symbol being analyzed
        gpi = 0                                                     ' array position of greatest % increase
        gpd = 0                                                     ' array position of greatest % decrease
        gtv = 0                                                     ' array position of greatest total volume
        notPublic = False                                           ' bool to check for IPO's
        IPOs = 0                                                    ' counter for how many IPO's this year
        
        '''''''''''''''''''''''''''''''''''''''''''''Need to add checks for failed matches for tX, vX, oX and cX in the future
    
        '''''''''''''''''''''''''''''''''''''''''''''Identify how many different ticker symbols there are
        Do While (Cells(dY, tX) <> "")
            If (Cells(dY, tX) <> tCur) Then
                size = size + 1
                tCur = Cells(dY, tX)
            End If
            dY = dY + 1
        Loop
        
        If (dY = 2) Then                        'End routine if there is no appropriate data
            MsgBox ("No data to operate on!")
            Exit Sub
        Else
            MsgBox ("Found " + Str(size) + " different tickers!")
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''Find space to output results
        Do While (Cells(1, pX) <> "")
            pX = pX + 1
        Loop
        pX = pX + 1
        Cells(1, pX) = "Evaluation:"
        lAdr = Split(Cells(1, pX).Address(True, False), "$")(0)     'note output range column letters
        rAdr = Split(Cells(1, pX + 8).Address(True, False), "$")(0)
        pX = pX + 1                                                 'move remaining print past labels
        
        ReDim tSym(size)                            'Resize arrays appropriately
        ReDim tVol(size)
        ReDim tOpn(size)
        ReDim tCls(size)
        dY = 2                                      'Reset row tracker
        tCur = Cells(dY, tX)
        tSym(tPos) = tCur
        tVol(tPos) = 0
        tOpn(tPos) = Cells(dY, oX)
        
        If (tOpn(tPos) = 0) Then                    'Adjust IPO value if it hasn't started trading yet
            notPublic = True
        End If
        
        
        '''''''''''''''''''''''''''''''''''''''''''''Main evaluation loop
        Do While (tCur <> "")
            
            If (tCur = tSym(tPos)) Then             'Same ticker symbol
                
                tVol(tPos) = tVol(tPos) + Cells(dY, vX) 'Aggregate volume
                
                If (notPublic) Then                     'Check for start of trading
                    If (Cells(dY, oX) > 0) Then         'Could have used AND here, but it doesn't feel right
                        tOpn(tPos) = Cells(dY, oX)      ' nested structure would make more sense for handling
                        notPublic = False               ' more cases in case the data was formated differently?
                        IPOs = IPOs + 1
                    End If
                End If
            End If
            
            '''''''''''''''''''''''''''''''''''''''''Print results if mismatch or out of tickers
            If ((tCur <> tSym(tPos)) Or (Cells(dY + 1, tX) = "")) Then
                
                If (Cells(dY + 1, tX) = "") Then        'Closing position if last item in list
                    tCls(tPos) = Cells(dY, cX)
                Else
                    tCls(tPos) = Cells(dY - 1, cX)      'Closing position if new ticker symbol
                End If
                
                '''''''''''''''''''''''''''''''''''''Report ticker info before moving on
                Cells(tPos + 2, pX) = tSym(tPos)
                Cells(tPos + 2, pX + 1) = tCls(tPos) - tOpn(tPos)
                
                '''''''''''''''''''''''''''''''''''''Format display of annual change
                Cells(tPos + 2, pX + 1).NumberFormat = "0.00"
                If (Cells(tPos + 2, pX + 1) > 0) Then
                    Cells(tPos + 2, pX + 1).Interior.Color = 10092441
                ElseIf (Cells(tPos + 2, pX + 1) < 0) Then
                    Cells(tPos + 2, pX + 1).Interior.Color = 10066431
                End If
                
                '''''''''''''''''''''''''''''''''''''Check for untraded stock errors
                If (tOpn(tPos) > 0) Then
                    Cells(tPos + 2, pX + 2) = Cells(tPos + 2, pX + 1) / tOpn(tPos)
                Else
                    MsgBox ("Ticker symbol '" + tSym(tPos) + "' didn't start trading in " + ActiveSheet.Name + "!")
                    Cells(tPos + 2, pX + 2) = 0
                End If
                
                '''''''''''''''''''''''''''''''''''''Format for percentage
                Cells(tPos + 2, pX + 2).NumberFormat = "0.00%"
                Cells(tPos + 2, pX + 3) = tVol(tPos)
                
                '''''''''''''''''''''''''''''''''''''Check for Greatest Total Volume
                If (tVol(tPos) > tVol(gtv)) Then
                    gtv = tPos
                    gtvT = tSym(tPos)
                ElseIf ((tVol(tPos) = tVol(gtv)) And (tPos > 0)) Then   'in case of ties
                    gtvT = gtvT + ", " + tSym(tPos)
                End If
                
                '''''''''''''''''''''''''''''''''''''Check for Greatest Percentage Increase
                If (Cells(tPos + 2, pX + 2) > Cells(gpi + 2, pX + 2)) Then
                    gpi = tPos
                    gpiT = tSym(tPos)
                ElseIf ((tPos > 0) And (Cells(tPos + 2, pX + 2) = Cells(gpi + 2, pX + 2))) Then 'in case of ties
                    gpiT = gpiT + ", " + tSym(tPos)
                End If
                
                '''''''''''''''''''''''''''''''''''''Check for Greatest Percentage Decrease
                If (Cells(tPos + 2, pX + 2) < Cells(gpd + 2, pX + 2)) Then
                    gpd = tPos
                    gpdT = tSym(tPos)
                ElseIf ((tPos > 0) And (Cells(tPos + 2, pX + 2) = Cells(gpd + 2, pX + 2))) Then 'in case of ties
                    gpdT = gpdT + ", " + tSym(tPos)
                End If
                
                tPos = tPos + 1
                tSym(tPos) = tCur
                tVol(tPos) = Cells(dY, vX)
                tOpn(tPos) = Cells(dY, oX)
                
                If (tOpn(tPos) = 0) Then        'Adjust status if the stock hasn't started trading yet
                    notPublic = True
                End If
            End If
            
            dY = dY + 1                         'Iterate
            tCur = Cells(dY, tX)
        Loop
        
        '''''''''''''''''''''''''''''''''''''''''''''Print Labels and "Greatest" results
        Cells(1, pX) = "Ticker"
        Cells(1, pX + 1) = "Yearly Change"
        Cells(1, pX + 2) = "Percent Change"
        Cells(1, pX + 3) = "Total Stock Volume"
        Cells(1, pX + 5) = "Greatest Values"
        Cells(2, pX + 5) = "% Increase"
        Cells(3, pX + 5) = "% Decrease"
        Cells(4, pX + 5) = "Total Volume"
        Cells(1, pX + 6) = "Ticker(s)"
        Cells(1, pX + 7) = "Value"
        
        Cells(2, pX + 6) = gpiT                     'Greatest results
        Cells(3, pX + 6) = gpdT
        Cells(4, pX + 6) = gtvT
        Cells(2, pX + 7) = Cells(gpi + 2, pX + 2)
        Cells(2, pX + 7).NumberFormat = "0.00%"
        Cells(3, pX + 7) = Cells(gpd + 2, pX + 2)
        Cells(3, pX + 7).NumberFormat = "0.00%"
        Cells(4, pX + 7) = tVol(gtv)
        
        '''''''''''''''''''''''''''''''''''''''''''''Format the columns for visibility
        Columns(lAdr + ":" + rAdr).EntireColumn.AutoFit
        Columns(lAdr + ":" + rAdr).EntireColumn.HorizontalAlignment = xlRight
        Range(lAdr + "1:" + rAdr + "1").Select
        Selection.Font.Bold = True
        Selection.HorizontalAlignment = xlCenter
        
        'MsgBox ("There were " + Str(IPOs) + " mid-year IPO's in " + ActiveSheet.Name + "!")
    Next n
End Sub

