Attribute VB_Name = "Module7"
Sub AddPercentageSymbolToKColumn()
    Dim ws As Worksheet
    Dim lastRowK As Long
    Dim i As Long

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in column K for each sheet
        lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).row

        ' Format cells in column K as percentages
        ws.Range("K2:K" & lastRowK).NumberFormat = "0.00%"

    Next ws

    MsgBox "Percentage symbol added to column K for all sheets."
End Sub

