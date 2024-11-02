Attribute VB_Name = "Module1"
Sub Stocked_Up()

Dim ws As Worksheet

    'Ticker Variables,
    Dim ticker As String
    Dim nextTicker As String
    
    ' Volume and Added Volume amounts will have huge numbers,
    ' Solved: use longlong(64-bit) not long(32-bit), no overflow
    Dim volume As LongLong
    Dim volumeAll As LongLong
    
    ' Variables to Move Down Rows and add to table
    Dim i As Long
    Dim tableRow As Long
    Dim lastRow As Long

    ' Variables For Quarterly Changes
    Dim openAmount As Double
    Dim closingAmount As Double
    Dim change As Double
    Dim pctChange As Double

    For Each ws In ThisWorkbook.Worksheets
        ' Organize the two tables that are being made
        
        ' First Table: Quarterly Changes , Percentage Change and Total Stock Volume portion
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Quarterly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        
        ' Second Table: Stock with Greatest % Increase, Greatest % Decrease,
        ' and Greatest Total Volume
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' Starts volume total amount, 0 is used as the natural point of counting value
        volumeAll = 0
        ' Starts the stock openning amount at cell C2, as that is the first value row for open
        openAmount = ws.Cells(2, 3).Value
        ' Like openAmount, row 2 is the first row with data, so we want to start counting from row 2
        tableRow = 2
        
        ' Find the last row with data in first column (Ticker)
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Allows the i count to stop at the last row (code made above) instead of going infinitely
        For i = 2 To lastRow
            ' Check Ticker and Volume amounts from original table
            ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            ' The code below allows us to use i to move to the next ticker row and
            ' Additionally will play a part in adding up all volumes of same tickers for future code
            nextTicker = ws.Cells(i + 1, 1).Value

            ' Check if next ticker is not the same as current ticker
            If (ticker <> nextTicker) Then
                ' add up to the total and close out the rest of the operations
                volumeAll = volumeAll + volume

                ' find amount and percentage change per ticker
                
                closingAmount = ws.Cells(i, 6).Value
                change = closingAmount - openAmount
                ' No zero values for open or closing stocks, no need to check for divide by 0
                ' ^Useful if worksheet needs values to be updated/changed (Not applicable here)
                pctChange = change / openAmount

                ' transfer to first table from J -> M
                ws.Cells(tableRow, 13).Value = volumeAll
                ' FormatPercent() adds formatting "%" to value
                ws.Cells(tableRow, 12).Value = FormatPercent(pctChange)
                ws.Cells(tableRow, 11).Value = change
                ws.Cells(tableRow, 10).Value = ticker

                ' Color Code Conditional Formating(positive = green, negative = red)
                If (change > 0) Then
                    ' Makes cell interior fill with green
                    ws.Cells(tableRow, 11).Interior.ColorIndex = 4
                ElseIf (change < 0) Then
                    ' Makes cell interior fill with red
                    ws.Cells(tableRow, 11).Interior.ColorIndex = 3
                Else
                    ' All cells returning with 0 are blank/white
                End If

                ' reset total volume for next ticker symbol and move down one in first table
                volumeAll = 0
                tableRow = tableRow + 1
                openAmount = ws.Cells(i + 1, 3).Value ' the open price of the NEXT stock (tricky gotcha)
            Else
                ' add total and loop back for checking next row for the ticker being the same
                volumeAll = volumeAll + volume
            End If
        Next i

        ' Second Loop for Second Leaderboard
        Dim maxPrice As Double
        Dim minPrice As Double
        Dim maxVolume As LongLong ' The highest total volumes will need 64-bit capacity
        Dim maxPriceTicker As String
        Dim minPriceTicker As String
        Dim maxVolumeTicker As String
        Dim j As Integer

        ' Start at first row of first made table to begin checking rows
        ' Will Find highest percentage change from "Percentage Change"
        maxPrice = ws.Cells(2, 12).Value
        ' Will Find most negative percentage change from "Percentage Change"
        minPrice = ws.Cells(2, 12).Value
        ' Will find highest volume from "Total Stock Volume"
        maxVolume = ws.Cells(2, 13).Value
        ' Returns ticker associated with maxPrice
        maxPriceTicker = ws.Cells(2, 10).Value
        ' Returns ticker associated with minPrice
        minPriceTicker = ws.Cells(2, 10).Value
        ' Returns ticker associated with maxVolumeTicker
        maxVolumeTicker = ws.Cells(2, 10).Value

        'Overwrite maxPrice with any value that is higher than what is currently in maxPrice
        For j = 2 To tableRow 'Start at second row
            ' Checks whether cell value is higher than stored value in maxPrice
            If (ws.Cells(j, 12).Value > maxPrice) Then
                ' If true, replace maxPrice with new higher value and correct ticker
                maxPrice = ws.Cells(j, 11).Value
                maxPriceTicker = ws.Cells(j, 10).Value
            End If
            ' As above, does the opposite, where it checks and replaces with lower prices
            If (Cells(j, 11).Value < minPrice) Then
                ' If true, replace minPrice with new higher value and correct ticker
                minPrice = ws.Cells(j, 12).Value
                minPriceTicker = ws.Cells(j, 10).Value
            End If

            If (Cells(j, 12).Value > maxVolume) Then
                ' If true, replace maxVolume with new higher value and correct ticker
                maxVolume = ws.Cells(j, 13).Value
                maxVolumeTicker = ws.Cells(j, 10).Value
            End If
        Next j

        ' Write second table
        ws.Range("P2").Value = maxPriceTicker
        ws.Range("P3").Value = minPriceTicker
        ws.Range("P4").Value = maxVolumeTicker

        ws.Range("Q2").Value = FormatPercent(maxPrice)
        ws.Range("Q3").Value = FormatPercent(minPrice)
        ws.Range("Q4").Value = maxVolume
    'Goes to next Quarterly worksheet
    Next ws
End Sub
'Reset for testing just in case anything messes up, and can reuse code for new values (Especially if there are less tickers)
Sub reset()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ' Delete columns I through Q, resets to original state
        ws.Range("I:Q").Delete
    'Goes to next Quarterly worksheet
    Next ws

End Sub
