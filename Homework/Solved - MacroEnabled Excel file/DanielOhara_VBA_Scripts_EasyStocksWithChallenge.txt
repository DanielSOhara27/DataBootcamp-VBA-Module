Attribute VB_Name = "Module1"
Sub EasyStocksWithChallenge():

    For Each ws In Worksheets
    CurrCrow = 1
    
        'Determine the last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Determine the last Column of data and set columns for Results
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        'lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
        CurrCColumn = lastCol + 2
        
        'Setting  the headers for column results
        ws.Cells(1, CurrCColumn).Value = "Ticker"
        ws.Cells(1, CurrCColumn + 1).Value = "Total Stock Volume"
        
        'sorting contents of each worksheet based on Column A
        ws.Columns("A:G").Sort key1:=ws.Range("A:A"), order1:=xlAscending, Header:=xlYes
        
        'Now that Columns are sorted, it is you can begin adding the total of each stock to the right
        For i = 2 To LastRow
        
            'If there current Cell value at Column A is different from the one above
            'then increase CurrCrow counter by 1 and then create a new placeholder
            'for the Total Stock Volume
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                CurrCrow = CurrCrow + 1
                ws.Cells(CurrCrow, CurrCColumn).Value = ws.Cells(i, 1).Value
                ws.Cells(CurrCrow, CurrCColumn + 1).Value = ws.Cells(i, lastCol).Value
            
            'Else it is the same stock as the previous iteration and thus you only need to add the
            'The Total Stock Volume by the new increment.
            Else
                ws.Cells(CurrCrow, CurrCColumn + 1) = ws.Cells(CurrCrow, CurrCColumn + 1).Value + ws.Cells(i, lastCol).Value
            End If
        Next i
        'MsgBox ("Sheet:" + ws.Name + " Ticker Name: " + ws.Cells(2, 1).Value)
    Next ws

End Sub

