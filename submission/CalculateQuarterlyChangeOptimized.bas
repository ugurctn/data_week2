Attribute VB_Name = "Module3"
Sub CalculateQuarterlyChangeOptimized()
    Dim ws As Worksheet
    Dim lastRowI As Long, lastRowA As Long
    Dim ticker As String
    Dim firstOpenData As Double, lastCloseData As Double
    Dim tickerDict As Object
    Dim tickerKey As String
    Dim i As Long, row As Long
    Dim dataA As Variant, dataC As Variant, dataF As Variant
    Dim outputArray() As Double
    Dim tickerData As Variant
    
    ' Attempt to use Dictionary; fallback to Collection if unavailable
    On Error Resume Next
    Set tickerDict = CreateObject("Scripting.Dictionary")
    If tickerDict Is Nothing Then
        Set tickerDict = New Collection
    End If
    On Error GoTo 0

    ' Turn off screen updating and calculations to speed up the process
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Clear any previous dictionary entries
        If TypeName(tickerDict) = "Dictionary" Then
            tickerDict.RemoveAll
        Else
            Set tickerDict = New Collection
        End If

        ' Find the last rows in columns A and I
        lastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
        lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        ' Load columns A, C, and F into arrays for faster processing
        dataA = ws.Range("A1:A" & lastRowA).Value
        dataC = ws.Range("C1:C" & lastRowA).Value
        dataF = ws.Range("F1:F" & lastRowA).Value

        ' Scan column A once to find the first and last row for each ticker
        For row = 2 To lastRowA
            ticker = dataA(row, 1)
            If ticker <> "" Then
                If TypeName(tickerDict) = "Dictionary" Then
                    ' Dictionary approach for fast key handling
                    If Not tickerDict.exists(ticker) Then
                        tickerDict.Add ticker, Array(row, row)
                    Else
                        tickerData = tickerDict(ticker)
                        tickerDict(ticker) = Array(tickerData(0), row)
                    End If
                Else
                    ' Collection approach for compatibility
                    On Error Resume Next
                    tickerDict.Add Array(row, row), ticker
                    If Err.Number = 457 Then
                        tickerData = tickerDict(ticker)
                        tickerDict.Remove ticker
                        tickerDict.Add Array(tickerData(0), row), ticker
                    End If
                    On Error GoTo 0
                End If
            End If
        Next row

        ' Prepare an output array for column J
        ReDim outputArray(1 To lastRowI - 1, 1 To 1)
        
        ' Calculate quarterly change for each ticker in column I
        For i = 2 To lastRowI ' Start from I2, assuming I1 is the header
            ticker = ws.Cells(i, "I").Value
            If TypeName(tickerDict) = "Dictionary" Then
                If tickerDict.exists(ticker) Then
                    tickerData = tickerDict(ticker)
                    firstOpenData = dataC(tickerData(0), 1)
                    lastCloseData = dataF(tickerData(1), 1)
                    outputArray(i - 1, 1) = lastCloseData - firstOpenData
                End If
            Else
                On Error Resume Next
                tickerData = tickerDict(ticker)
                On Error GoTo 0
                If Not IsEmpty(tickerData) Then
                    firstOpenData = dataC(tickerData(0), 1)
                    lastCloseData = dataF(tickerData(1), 1)
                    outputArray(i - 1, 1) = lastCloseData - firstOpenData
                End If
            End If
        Next i

        ' Write the output array to column J
        ws.Range("J2").Resize(UBound(outputArray, 1), 1).Value = outputArray

        ' Apply conditional formatting to column J
        With ws.Range("J2:J" & lastRowI)
            .FormatConditions.Delete ' Clear existing formatting
            ' Add conditional format for negative values
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0) ' Red
            ' Add conditional format for positive values
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0) ' Green
        End With
    Next ws

    ' Turn screen updating and calculations back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Quarterly changes calculated and added to column J for all sheets, with color coding applied."
End Sub

