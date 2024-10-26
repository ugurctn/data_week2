Attribute VB_Name = "Module5"
Sub CalculateTotalVolume()
    Dim ws As Worksheet
    Dim lastRowI As Long, lastRowA As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim i As Long, row As Long
    Dim dataA As Variant, dataG As Variant
    
    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last rows in columns A and I
        lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        ' Load columns A and G into arrays for faster processing
        dataA = ws.Range("A1:A" & lastRowA).Value
        dataG = ws.Range("G1:G" & lastRowA).Value
        
        ' Loop through each ticker in column I
        For i = 2 To lastRowI ' Start from I2, assuming I1 is the header
            ticker = ws.Cells(i, "I").Value
            totalVolume = 0
            
            ' Sum the volume for each occurrence of the ticker in column A
            For row = 2 To lastRowA
                If dataA(row, 1) = ticker Then
                    totalVolume = totalVolume + dataG(row, 1)
                End If
            Next row
            
            ' Write the total volume to column L in the same row as the ticker
            ws.Cells(i, "L").Value = totalVolume
        Next i
    Next ws
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True

    MsgBox "Total stock volume calculated and added to column L for all sheets."
End Sub

