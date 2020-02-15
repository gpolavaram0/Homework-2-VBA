Attribute VB_Name = "Module1"
Sub StockInfo():
    
    'Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'Define var types
    
    Dim SheetRowCount As Long
    Dim SheetRowColumn As Long
    
    Dim FirstTicker As String
    Dim CurrentTicker As String
      
    Dim FirstTickerRow As Long
    Dim CurrentTickerRowEnd As Long
    Dim CurrentTickerRowFirst As Long
    Dim CurrentPercChange As Double
     
    Dim DataRowCount As Long
        
    '---------------------------------
    'Label Data Columns
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    
    '---------------------------------
    
    SheetRowCount = ws.Cells(Rows.Count, 2).End(xlUp).Row
    SheetRowColumn = ws.Cells(2, Columns.Count).End(xlToLeft).Column
    
    DataRowCount = 2
    
    CurrentTicker = ws.Cells(2, "A")
    CurrentTickerRowFirst = 2
    
    '-------------------------------------
    
    
   ' For c = 1 To SheetRowColumn
    
        For r = 2 To SheetRowCount + 1
        
           If ws.Cells(r, 1) <> CurrentTicker Or r = SheetRowCount + 1 Then
           
                
                
           
                'CurrentTickerRowFirst
                
                CurrentTickerRowEnd = ws.Cells(r - 1, 1).Row
                
                VolumeSum = WorksheetFunction.Sum(Range(ws.Cells(CurrentTickerRowFirst, "G"), ws.Cells(CurrentTickerRowEnd, "G")))
                
                    If ws.Cells(CurrentTickerRowFirst, "C") <> 0 Then
                
                        CurrentPercChange = 100 * (ws.Cells(CurrentTickerRowEnd, "F") - ws.Cells(CurrentTickerRowFirst, "C")) / ws.Cells(CurrentTickerRowFirst, "C")
                
                    Else
                        
                        CurrentPercChange = 0
                    
                    End If
                
                
                TickerYearlyChange = ws.Cells(CurrentTickerRowEnd, "F") - ws.Cells(CurrentTickerRowFirst, "C")
                'MsgBox (CurrentPercChange)
                'MsgBox (CurrentTickerRowFirst)
                'MsgBox (CurrentTickerRowEnd)
                ws.Cells(DataRowCount, "L").Value = VolumeSum
                ws.Cells(DataRowCount, "K").Value = CurrentPercChange
                ws.Cells(DataRowCount, "J").Value = TickerYearlyChange
                
                If TickerYearlyChange < 0 Then
                
                    ws.Cells(DataRowCount, "J").Interior.ColorIndex = 3
                    
                Else
                    
                    ws.Cells(DataRowCount, "J").Interior.ColorIndex = 4
                
                End If
                
                
                
                
                
                ws.Cells(DataRowCount, "I").Value = CurrentTicker
                DataRowCount = DataRowCount + 1
                CurrentTicker = ws.Cells(r, 1)
                CurrentTickerRowFirst = ws.Cells(r, 1).Row
                
                CurrentTickerRowFirst = ws.Cells(r, 1).Row
                'MsgBox (ws.Cells(r, 1).Row)
                'MsgBox (ws.Cells(r - 1, 1) + ":First")
           
           End If
        
        Next r
    
    For n = 1 To Application.Sheets.Count
    Worksheets(n).Activate
    
    
    GreatestPercInc = WorksheetFunction.Max(Columns("K"))
    GreatestPercDec = ws.Application.WorksheetFunction.Min(Columns("K"))
    GreatestTotVolume = ws.Application.WorksheetFunction.Max(Columns("L"))
    
      
    'ws.Cells(2, "Q") = GreatestPercInc
    'ws.Cells(3, "Q") = GreatestPercDec
    'ws.Cells(4, "Q") = GreatestTotVolume
           
    For r = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(r, 11) = GreatestPercInc Then
        
            ws.Cells(2, "P") = ws.Cells(r, 9)
            ws.Cells(2, "Q") = ws.Cells(r, 11)
        
        End If
        
    Next r
    
    For r = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
        If ws.Cells(r, 11) = GreatestPercDec Then
        
            ws.Cells(3, "P") = ws.Cells(r, 9)
            ws.Cells(3, "Q") = ws.Cells(r, 11)
        
        End If
        
    Next r
    
    For r = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
        If ws.Cells(r, 12) = GreatestTotVolume Then
        
            ws.Cells(4, "P") = ws.Cells(r, 9)
            ws.Cells(4, "Q") = ws.Cells(r, 12)
        
        End If
        
    Next r
    
    Next n
    
    Dim lRow As Long

    lRow = ws.Cells(6, 1).Row
    
    'MsgBox (TempRowCount)
    
    ws.Cells(1, 1).Interior.ColorIndex = 0
           
    Next

End Sub





