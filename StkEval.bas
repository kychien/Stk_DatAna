Attribute VB_Name = "StkDatAna"
Sub StkEval():
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Script to cycle through several excel sheets to gather stock
    '   volume over the years.
    '
    '   2018 08 10 - Basic implementation for 1 sheet of data to analyze total volume
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim tSym() As String                        'Array to hold ticker symbols for aggregating data over years
    Dim tVol() As Double                        'Array to track corresponding volume
    Dim tx, vx, px, ty, py, size, tPos As Integer
    Dim tCur As String
    
    '''''''''''''''''''''''''''''''''''''''''''''Initialize excel coordiantes...
    tx = 1                                      ' ticker column
    ty = 2                                      ' ticker row
    vx = 7                                      ' volume column
    px = 10                                     ' printout column
    py = 2                                      ' printout row
    size = 0                                    ' number of unique ticker symbols
    tPos = 0                                    ' array position tracker
    tCur = Cells(ty, tx)                        ' current ticker symbol being analyzed
    
    '''''''''''''''''''''''''''''''''''''''''''''Identify how many different ticker symbols there are
    Do While (Cells(ty, tx) <> "")              'While the next cell has data...
        If (Cells(ty, tx) <> tCur) Then         'If the ticker is unique...
            size = size + 1
        End If
        ty = ty + 1
    Loop
    
    If (size < 1) Then                          'End routine if there is no appropriate data
        MsgBox ("No data to operate on!")
        Exit Sub
    End If
    
    'MsgBox ("Found " + Str(size) + " different tickers!")
    
    ReDim tSym(size)                            'Resize arrays appropriately
    ReDim tVol(size)
    tx = 1                                      'Reset excel coordinate vars
    ty = 2
    tCur = Cells(ty, tx)
    tSym(tPos) = tCur
    tVol(tPos) = 0
    
    '''''''''''''''''''''''''''''''''''''''''''''Collect volume data
    Do While (tCur <> "")                       'While the next cell has data...
        
        If (tCur = tSym(tPos)) Then             'If the ticker is the same...
            tVol(tPos) = tVol(tPos) + Cells(ty, vx)
        Else
            Cells(tPos + 2, px) = tSym(tPos)
            'Cells(tPos + 2, px + 3) = tVol(tPos)
            Cells(tPos + 2, px + 1) = tVol(tPos)
            tPos = tPos + 1
            tSym(tPos) = tCur
            tVol(tPos) = Cells(ty, vx)
        End If
        
        ty = ty + 1                       'Iterate
        tCur = Cells(ty, tx)
    Loop
    
    '''''''''''''''''''''''''''''''''''''''''''''Format Results and Labels on corresponding sheet
    Cells(1, px) = "Ticker"
    'Cells(1, px + 1) = "Yearly Change"
    'Cells(1, px + 2) = "Percent Change"
    'Cells(1, px + 3) = "Total Stock Volume"
    Cells(1, px + 1) = "Total Stock Volume"
    Columns("J:M").EntireColumn.AutoFit
    Columns("J:M").EntireColumn.HorizontalAlignment = xlRight
    Range("I1:M1").Select
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
    Selection.Font.Bold = True
    
End Sub

