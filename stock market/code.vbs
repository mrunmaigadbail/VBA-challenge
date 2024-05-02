Option Explicit
Sub stock()
    'variable declaration
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim k As Integer
    Dim open_count As Integer
    Dim i As Integer
    Dim ticker As String
    Dim open_val As Double
    Dim close_val As Double
    Dim Quarterly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    For Each ws In Worksheets ' for all worksheets
    ' add column names
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Quarterly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        'autofit colums from a to z
        ws.Columns("A:Z").AutoFit
        
        'finding last row of sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '% format
        ws.Range("K:K").NumberFormat = "0.00%"
        k = 2
        Dim volume As Double
        open_count = 2
        For i = 2 To LastRow
            'data retrival
            volume = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
                       
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
        
                'data retrival
                open_val = ws.Cells(open_count, 3).Value
                close_val = ws.Cells(i, 6).Value
                
                'to check if we are getting correct values
                'ws.Cells(k, 15).Value = open_val
                'ws.Cells(k, 16).Value = close_val
                
                'formulas for calculations
                Quarterly_change = close_val - open_val
                percent_change = Quarterly_change / open_val
                total_volume = total_volume + volume
                
                'assigning values to columns
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = Quarterly_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = total_volume
                
                ' reset
                total_volume = 0
                k = k + 1
                open_count = i + 1
            Else
                ' we just add to the total
                total_volume = total_volume + volume
            End If
        Next i
        
        
        'conditional formatting for quarterly change and percent change
        Dim LastRowChange As Long
        LastRowChange = ws.Cells(Rows.Count, 10).End(xlUp).Row
        'For i = 2 To LastRowChange
          '  If ws.Cells(i, 10).Value > 0 Then
           '     ws.Cells(i, 10).Interior.ColorIndex = 4
            '    ws.Cells(i, 11).Interior.ColorIndex = 4
            'End If
            'If ws.Cells(i, 10).Value < 0 Then
            '    ws.Cells(i, 10).Interior.ColorIndex = 3
             '   ws.Cells(i, 11).Interior.ColorIndex = 3
            'End If
            
        'Next i
       ' For i = 2 To LastRowChange
            Dim rng As Range
            Set rng = ws.Range("J2:k" & LastRowChange)
            With rng
                .FormatConditions.Delete
            
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
                .FormatConditions(1).Interior.ColorIndex = 4
                
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=0
                .FormatConditions(2).Interior.ColorIndex = 3
                
            End With
       ' Next i
    
       
       
       ' inserting greatest increase, decrease, volume
       
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
        ws.Range("o2").Value = "Greatest % increase"
        ws.Range("o3").Value = "greatest % decrease"
        ws.Range("o4").Value = "greatest total volume"
        
        'code for greatest % increase
        Dim maxIncrease As Double
        Dim maxTickerI As String
        
        maxIncrease = 0
        For i = 2 To LastRowChange
            If (ws.Cells(i, 11).Value > maxIncrease) Then
                maxIncrease = ws.Cells(i, 11).Value
                maxTickerI = ws.Cells(i, 9).Value
    
            End If
            
        Next i
        ws.Range("q2").NumberFormat = "0.00%"
        ws.Range("q2") = maxIncrease
        ws.Range("p2") = maxTickerI
        
        
        'code for greatest % Decrease
        Dim maxDecrease As Double
        Dim maxTickerD As String
        
        maxDecrease = 0
        For i = 2 To LastRowChange
            If (ws.Cells(i, 11).Value < maxDecrease) Then
                maxDecrease = ws.Cells(i, 11).Value
                maxTickerD = ws.Cells(i, 9).Value
    
            End If
            
        Next i
        ws.Range("q3").NumberFormat = "0.00%"
        ws.Range("q3") = maxDecrease
        ws.Range("p3") = maxTickerD
        'code for greatest total volume
        Dim maxVolume As Double
        Dim maxTickerV As String
        
        maxVolume = 0
        For i = 2 To LastRowChange
            If (ws.Cells(i, 12).Value > maxVolume) Then
                maxVolume = ws.Cells(i, 12).Value
                maxTickerV = ws.Cells(i, 9).Value
    
            End If
            
        Next i
        
        ws.Range("q4") = maxVolume
        ws.Range("p4") = maxTickerV
    Next ws
End Sub

