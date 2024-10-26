Attribute VB_Name = "Module4"
Sub CalculatePercentageChange()
    Dim ws As Worksheet
    Dim lastRowI As Long, lastRowA As Long
    Dim ticker As String
    Dim QC As Double ' Quarterly Change from column J
    Dim FOD As Double ' First Open Data from column C
    Dim percentageChange As Double
    Dim i As Long, row As Long
    Dim dataA As Variant, dataC As Variant
    Dim foundRow As Long
    
    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last rows in columns A and I
        lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        ' Load columns A and C into arrays for faster processing
        dataA = ws.Range("A1:A" & lastRowA).Value
        dataC = ws.Range("C1:C" & lastRowA).Value
        
        ' Loop through each ticker in column I
        For i = 2 To lastRowI ' Start from I2, assuming I1 is the header
            ticker = ws.Cells(i, "I").Value
            
            ' Get the Quarterly Change (QC) from column J
            QC = ws.Cells(i, "J").Value
            
            ' Find the first occurrence of the ticker in column A to get FOD
            foundRow = 0
            For row = 2 To lastRowA
                If dataA(row, 1) = ticker Then
                    FOD = dataC(row, 1)
                    foundRow = row
                    Exit For
                End If
            Next row
            
            ' Check if FOD was found and avoid division by zero
            If foundRow > 0 And FOD <> 0 Then
                ' Calculate percentage change
                percentageChange = (QC / FOD) * 100
            Else
                ' If FOD not found or is zero, set percentage change to zero
                percentageChange = 0
            End If
            
            ' Write the result to column K in the same row as the ticker
            ws.Cells(i, "K").Value = percentageChange
        Next i
    Next ws
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True

    MsgBox "Percentage change calculated and added to column K for all sheets."
End Sub

