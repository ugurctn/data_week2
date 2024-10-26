Attribute VB_Name = "Module2"
Sub AddHeadersToAllSheets()
    Dim ws As Worksheet
    
    ' Turn off screen updating and calculations to speed up the process
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Add headers in specified columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Headers for the "Greatest" analysis
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    Next ws
    
    ' Turn screen updating and calculations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Headers have been added to all sheets."
End Sub

