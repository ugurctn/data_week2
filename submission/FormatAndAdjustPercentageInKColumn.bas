Attribute VB_Name = "Module6"
Sub FormatAndAdjustPercentageInKColumn()
    Dim ws As Worksheet
    Dim lastRowK As Long
    Dim i As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in column K for each sheet
        lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).row

        ' Loop through each cell in column K to adjust and format as percentage
        For i = 2 To lastRowK ' Start from K2, assuming K1 is the header
            ws.Cells(i, "K").Value = ws.Cells(i, "K").Value / 100
        Next i

        ' Format cells in column K as percentages
        ws.Range("K2:K" & lastRowK).NumberFormat = "0.00%"
    Next ws

    MsgBox "Values in column K adjusted and formatted as percentages for all sheets."
End Sub

