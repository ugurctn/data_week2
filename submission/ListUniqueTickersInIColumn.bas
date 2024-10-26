Attribute VB_Name = "Module1"
Sub ListUniqueTickersInIColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As Variant
    Dim uniqueTickers As New Collection
    Dim outputArray() As String
    Dim i As Long, row As Long
    Dim exists As Boolean

    ' Turn off screen updating and calculations to speed up the process
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        ' Loop through column A and add unique tickers to the collection
        For row = 2 To lastRow ' Start from row 2, assuming row 1 is a header
            ticker = CStr(ws.Cells(row, "A").Value)
            If ticker <> "" Then
                exists = False
                ' Check if ticker already exists in the collection
                On Error Resume Next
                exists = (uniqueTickers(ticker) <> "")
                On Error GoTo 0
                ' Add ticker if it does not already exist
                If Not exists Then uniqueTickers.Add ticker, ticker
            End If
        Next row
    Next ws

    ' Prepare an array from the unique tickers collection
    ReDim outputArray(1 To uniqueTickers.Count)
    For i = 1 To uniqueTickers.Count
        outputArray(i) = uniqueTickers(i)
    Next i

    ' Insert unique tickers into column I starting from row 2 of each sheet
    For Each ws In ThisWorkbook.Sheets
        ws.Range("I2").Resize(UBound(outputArray), 1).Value = Application.Transpose(outputArray)
    Next ws

    ' Turn screen updating and calculations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Unique tickers listed in column I for all sheets."
End Sub

