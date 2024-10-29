Attribute VB_Name = "Module1"
Sub stonks()
    ' Variables
    Dim ws As Worksheet
    Dim stock As String
    Dim next_stock As String
    Dim volume As LongLong
    Dim volume_total As LongLong
    Dim i As Long
    Dim leaderboard_row As Long
    Dim lastRow As Long
    
    ' New variables
    Dim open_price As Double
    Dim closing_price As Double
    Dim change As Double
    Dim pct_change As Double
    
    For Each ws In ThisWorkbook.Worksheets
        ' Set Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

        ' Reset per stock
        volume_total = 0
        open_price = ws.Cells(2, 3).Value
        leaderboard_row = 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastRow
            ' Extract values from workbook
            stock = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            next_stock = ws.Cells(i + 1, 1).Value
    
            ' If statement
            If stock <> next_stock Then
                ' Add total
                volume_total = volume_total + volume
                
                ' Change logic
                closing_price = ws.Cells(i, 6).Value
                change = closing_price - open_price
                pct_change = change / open_price
    
                ' Write to leaderboard
                ws.Cells(leaderboard_row, 12).Value = volume_total
                ws.Cells(leaderboard_row, 11).Value = Format(pct_change, "0.00%")
                ws.Cells(leaderboard_row, 10).Value = change
                ws.Cells(leaderboard_row, 9).Value = stock
                
                ' Conditional Formatting
                If change > 0 Then
                    ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 4
                ElseIf change < 0 Then
                    ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 3
                End If
    
                ' Reset total
                volume_total = 0
                leaderboard_row = leaderboard_row + 1
                open_price = ws.Cells(i + 1, 3).Value ' Open price of the next stock
            Else
                ' Add total
                volume_total = volume_total + volume
            End If
        Next i
        
        ' Second Loop for Second Leaderboard
        Dim max_price As Double
        Dim min_price As Double
        Dim max_volume As LongLong
        Dim max_price_stock As String
        Dim min_price_stock As String
        Dim max_volume_stock As String
        Dim j As Long
        
        ' Initialize for comparison
        max_price = ws.Cells(2, 11).Value
        min_price = ws.Cells(2, 11).Value
        max_volume = ws.Cells(2, 12).Value
        max_price_stock = ws.Cells(2, 9).Value
        min_price_stock = ws.Cells(2, 9).Value
        max_volume_stock = ws.Cells(2, 9).Value
        
        For j = 2 To leaderboard_row
            ' Compare current row to the initial values
            If ws.Cells(j, 11).Value > max_price Then
                max_price = ws.Cells(j, 11).Value
                max_price_stock = ws.Cells(j, 9).Value
            End If
            
            If ws.Cells(j, 11).Value < min_price Then
                min_price = ws.Cells(j, 11).Value
                min_price_stock = ws.Cells(j, 9).Value
            End If
            
            If ws.Cells(j, 12).Value > max_volume Then
                max_volume = ws.Cells(j, 12).Value
                max_volume_stock = ws.Cells(j, 9).Value
            End If
        Next j
        
        ' Write out to Excel Workbook
        ws.Range("O2").Value = max_price_stock
        ws.Range("O3").Value = min_price_stock
        ws.Range("O4").Value = max_volume_stock
        
        ws.Range("P2").Value = max_price
        ws.Range("P3").Value = min_price
        ws.Range("P4").Value = max_volume
    Next ws
End Sub

Sub reset()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Delete columns I through P
        ws.Range("I:P").Delete
    Next ws
End Sub

