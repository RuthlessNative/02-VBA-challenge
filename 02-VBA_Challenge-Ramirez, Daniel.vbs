Sub VBA_challenge():

'Initiate variables
Dim ws As Integer
Dim lastdataRow As Long
Dim prevTicker As String
Dim newTicker As String
Dim totalTickers As Long
Dim volume As LongLong
Dim openVal As Double
Dim closeVal As Double
Dim yrCh As Double
Dim goatInc As Double
Dim goatIncTicker As String
Dim goatDec As Double
Dim goatDecTicker As String
Dim goatVol As Double
Dim goatVolTicker As String

'Assign global variables
ws = ActiveWorkbook.Worksheets.Count

'For Loop loops through each worksheet
For i = 1 To ws

    '(Re)assigns variables at the beginning of each worksheet
    lastdataRow = ActiveWorkbook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    prevTicker = ActiveWorkbook.Worksheets(i).Range("A2")
    newTicker = ActiveWorkbook.Worksheets(i).Range("A2").Value
    totalTickers = 0
    volume = 0
    openVal = ActiveWorkbook.Worksheets(i).Cells(2, 3).Value
    closeVal = 0
    yrCh = 0
    goatInc = 0
    goatDec = 0
    goatVol = 0

    'Print header titles at the beginning of each worksheet
    ActiveWorkbook.Worksheets(i).Cells(1, 9) = "Ticker"
    ActiveWorkbook.Worksheets(i).Cells(1, 10) = "Yearly Change"
    ActiveWorkbook.Worksheets(i).Cells(1, 11) = "Percent Change"
    ActiveWorkbook.Worksheets(i).Cells(1, 12) = "Total Stock Volume"
    ActiveWorkbook.Worksheets(i).Cells(2, 14) = "Greatest % Increase"
    ActiveWorkbook.Worksheets(i).Cells(3, 14) = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(i).Cells(4, 14) = "Greatest Total Volume"
    ActiveWorkbook.Worksheets(i).Cells(1, 15) = "Ticker"
    ActiveWorkbook.Worksheets(i).Cells(1, 16) = "Value"

    'Prints first ticker text value
    ActiveWorkbook.Worksheets(i).Cells(2, 9) = ActiveWorkbook.Worksheets(i).Range("A2").Value

    'For Loop goes 1 row past last row in data to capture previous row data:
    For j = 2 To (lastdataRow + 1)

        'Adds volume at every row iteration regardless of new ticker or not:
        volume = ActiveWorkbook.Worksheets(i).Cells(j, 7).Value + volume

        'If this new row's Ticker differs from the previous one then start the calculations, print, and reset variables
        If ActiveWorkbook.Worksheets(i).Cells(j, 1) <> prevTicker Then
            newTicker = ActiveWorkbook.Worksheets(i).Cells(j, 1)
            'Adds 1 to total count of tickers:
            totalTickers = totalTickers + 1
            'Prints next ticker text value accounting for header and first ticker text:
            ActiveWorkbook.Worksheets(i).Cells(totalTickers + 2, 9) = newTicker

            'Gets close value from previous row (last row of previous ticker):
            If j = 2 Then
                closeVal = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
            ElseIf j > 2 Then
                closeVal = ActiveWorkbook.Worksheets(i).Cells(j - 1, 6).Value
            End If
            'Calculates yearly change:
            yrCh = closeVal - openVal
            'Prints Yearly Change accounting for header:
            ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 10) = yrCh

            'Checks Yearly Change value and colors cells accordingly:
            If yrCh > 0 Then
                ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yrCh < 0 Then
                ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 10).Interior.ColorIndex = 3
            End If

            'Prints Percent Change accounting for header and avoids overflow when 0/0:
            If yrCh = 0 Or openVal = 0 Then
                ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 11) = 0 & "%"
            Else
                ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 11) = Round((yrCh / openVal) * 100, 2) & "%"
                'Stores yrCh and associated Ticker as either "Greatest % Increase" or "Greatest % Decrease" if applicable
                If (yrCh / openVal) > goatInc Then
                    goatIncTicker = prevTicker
                    goatInc = yrCh / openVal
                ElseIf (yrCh / openVal) < goatDec Then
                    goatDecTicker = prevTicker
                    goatDec = yrCh / openVal
                End If
            End If

            'Prints volume sum minus the newfound first newTicker volume accouting for header:
            ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 12) = volume - ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            'Stores that volume if it's the greatest so far
            If ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 12).Value > goatVol Then
                goatVolTicker = prevTicker
                goatVol = ActiveWorkbook.Worksheets(i).Cells(totalTickers + 1, 12).Value
            End If

            'Resets variables for the new ticker:
            prevTicker = newTicker
            volume = ActiveWorkbook.Worksheets(i).Cells(j, 7).Value
            openVal = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value

        End If

    Next j

    'Prints Greatest % Increase, % Decrease, and Volume before next Worksheet
    ActiveWorkbook.Worksheets(i).Cells(2, 15) = goatIncTicker
    ActiveWorkbook.Worksheets(i).Cells(2, 16) = Round(goatInc * 100, 2) & "%"
    ActiveWorkbook.Worksheets(i).Cells(3, 15) = goatDecTicker
    ActiveWorkbook.Worksheets(i).Cells(3, 16) = Round(goatDec * 100, 2) & "%"
    ActiveWorkbook.Worksheets(i).Cells(4, 15) = goatVolTicker
    ActiveWorkbook.Worksheets(i).Cells(4, 16) = goatVol

Next i

End Sub