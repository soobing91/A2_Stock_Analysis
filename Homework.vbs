Sub Homework()

    For Each ws In Worksheets

        'Declares dimensions for variables
        Dim Ticker As String
        Dim Opening, Closing, Change, Percentage, Total, MaxChange, MinChange, MaxTotal As Double

        Dim i As Long
        Dim LastRow As Long
        Dim RefRow, RefCol As Integer

        'Declares variables (and "constants") that will be used for loop
        Total = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        RefRow = 2
        RefCol = 9

        Set Refcell = ws.Cells(RefRow, RefCol)

        Refcell.Offset(-1, 0).Value = "Ticker"
        Refcell.Offset(-1, 1).Value = "Yearly Change"
        Refcell.Offset(-1, 2).Value = "Percent Change"
        Refcell.Offset(-1, 3).Value = "Total Stock Volume"
        Refcell.Offset(-1, 7).Value = "Ticker"
        Refcell.Offset(-1, 8).Value = "Value"
        Refcell.Offset(0, 6).Value = "Greatest % Increase"
        Refcell.Offset(1, 6).Value = "Greatest % Decrease"
        Refcell.Offset(2, 6).Value = "Greatest Total Volume"

        'Declaring the opening price at the beginning of each year for each Ticker
        Opening = ws.Cells(RefRow, 3).Value

        'Starting a "for Loop" statement
        For i = RefRow To LastRow
            Ticker = ws.Cells(i, 1).Value
            Closing = ws.Cells(i, 6).Value
            Change = Closing - Opening
            Total = Total + ws.Cells(i, 7).Value

            If Opening <> 0 Then
                Percentage = Change / Opening
            Else
                Percentage = 0
            End If

            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ws.Cells(RefRow, RefCol).Value = Ticker

                'Part 2: Calculating the difference between opening price and
                'closing price for each Ticker
                ws.Cells(RefRow, RefCol + 1).Value = Change
                ws.Cells(RefRow, RefCol + 2).Value = Percentage

                'Applying conditional formatting
                Set Dummy = ws.Cells(RefRow, RefCol + 1)
                    If Change > 0 Then
                        Dummy.Interior.ColorIndex = 4
                    ElseIf Change < 0 Then
                        Dummy.Interior.ColorIndex = 3
                    Else
                        Dummy.Interior.ColorIndex = 6
                    End If
                Set Dummy = Nothing

                'Part 1: Calculating subtotal for each Ticker
                ws.Cells(RefRow, RefCol + 3).Value = Total

                'Redeclaring the subtotal
                Total = 0

                'Moving the reference row to the next row
                RefRow = RefRow + 1

                'Moving on to the next opening price for the next Ticker
                Opening = ws.Cells(i + 1, 3).Value

            End If

        Next i

        'Part 3: Defining maximum and minimum
        MaxChange = WorksheetFunction.Max(ws.Range("K2:K" & (RefRow - 1)))
        MinChange = WorksheetFunction.Min(ws.Range("K2:K" & (RefRow - 1)))
        MaxTotal = WorksheetFunction.Max(ws.Range("L2:L" & (RefRow - 1)))

        ws.Cells(2, RefCol + 8).Value = MaxChange
        ws.Cells(3, RefCol + 8).Value = MinChange
        ws.Cells(4, RefCol + 8).Value = MaxTotal

        'Looking for Ticker that yields each max/min
        For i = 2 To (RefRow - 1)
            Set RefTik = ws.Cells(i, RefCol)
            If MaxChange = ws.Cells(i, RefCol + 2).Value Then
                ws.Cells(2, RefCol + 7).Value = RefTik.Value
            ElseIf MinChange = ws.Cells(i, RefCol + 2).Value Then
                ws.Cells(3, RefCol + 7).Value = RefTik.Value
            ElseIf MaxTotal = ws.Cells(i, RefCol + 3).Value Then
                ws.Cells(4, RefCol + 7).Value = RefTik.Value
            End If

            ws.Range("K2:K" & (RefRow - 1)).NumberFormat = "0.00%"
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

            Set RefTik = Nothing
        Next i

        ws.Columns("I:Q").AutoFit

    Next ws

    'Alert for completing a task
    MsgBox ("Task completed!")

End Sub