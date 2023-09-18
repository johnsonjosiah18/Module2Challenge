Sub MuSltstkyrdata()
    Dim ws As Worksheet
    Dim ticker As String
    Dim summarytable1 As Integer
    Dim lastrow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlychange As Double
    Dim percentageChange As Double
    Dim volume As Double
    Dim totalVolume As Double

    ' Variables for tracking greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

        summarytable1 = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Initialize totals for each ticker symbol
        totalVolume = 0

        ' Initialize greatest values
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Loop through each row in the current worksheet
        For i = 2 To lastrow

            ' Get ticker symbol for this row
            ticker = ws.Cells(i, 1).Value

            ' If next row has a different ticker symbol, calculate yearly change and percentage change for this ticker symbol
            If ws.Cells(i + 1, 1).Value <> ticker Then

                ' Get closing price for this ticker symbol
                closingPrice = ws.Cells(i, 6).Value

                ' Calculate yearly change and percentage change for this ticker symbol
                yearlychange = closingPrice - openingPrice
                percentageChange = (yearlychange / openingPrice)

                ' Print results to summary table for this ticker symbol
                ws.Range("I" & summarytable1).Value = ticker
                ws.Range("J" & summarytable1).Value = yearlychange
                
                    ' Format cells in column interior color to green if above 0 and red if below
                    If ws.Range("J" & summarytable1).Value >= 0 Then
                    ws.Range("J" & summarytable1).Interior.ColorIndex = 4

                    Else
                    ws.Range("J" & summarytable1).Interior.ColorIndex = 3

                    End If

            

            
                ws.Range("K" & summarytable1).Value = percentageChange
                    
                    'Format cells in column to percentage
                    ws.Range("K" & summarytable1).NumberFormat = "0.00%"
                
                    
                ws.Range("L" & summarytable1).Value = totalVolume

                ' Check if this ticker has the greatest increase, decrease, or volume so far
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    tickerGreatestIncrease = ticker
                ElseIf percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    tickerGreatestDecrease = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If

                ' Prepare for next ticker symbol
                summarytable1 = summarytable1 + 1

                ' Reset totals for next ticker symbol
                totalVolume = 0

            ElseIf ws.Cells(i - 1, 1).Value <> ticker Then

                ' Get opening price for this ticker symbol (assuming it's in column C)
                openingPrice = ws.Cells(i, 3).Value

            End If
            
            ' Add volume to total for this ticker symbol (assuming it's in column G)
            volume = ws.Cells(i, 7).Value
            totalVolume = totalVolume + volume

        Next i

        ' Print the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        ws.Range("P2").Value = "Ticker"
        ws.Range("Q2").Value = "Value"

        ws.Range("P3").Value = tickerGreatestIncrease
        ws.Range("Q3").Value = greatestIncrease

        ws.Range("P4").Value = tickerGreatestDecrease
        ws.Range("Q4").Value = greatestDecrease

        ws.Range("P5").Value = tickerGreatestVolume
        ws.Range("Q5").Value = greatestVolume

    Next ws

End Sub

