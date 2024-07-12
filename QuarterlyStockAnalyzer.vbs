Attribute VB_Name = "QuarterlyStockAnalyzer"
Sub QuarterlyStockAnalyzer()
    
'Declare the variables for the final summary metrics to be displayed
Dim Ticker As String
Dim qChange As Double
Dim pChange As Double
Dim totalVolume As Double

Dim greatestIncTicker As String
Dim greatestIncrease As Double
Dim greatestDecTicker As String
Dim greatestDecrease As Double
Dim greatestVolTicker As String
Dim greatestVolume As Double

'Declare intermediary variables to be used for calculating and formatting the final metrics to be displayed
Dim startPrice As Double
Dim endPrice As Double
Dim colorIndex As Integer
Dim summaryRow As Integer
Dim lastRow As Double

'Determine number of worksheets in the workbook
Dim worksheetsCount As Integer
worksheetsCount = Worksheets.Count


'Loop through each sheet in the workbook
Dim h As Integer
For h = 1 To worksheetsCount
    Worksheets(h).Activate


    'Determine the last row of a sheet's data set to be analyzed
    lastRow = Range("A2").End(xlDown).Row

    'Print the headings for the summary metrics tables and format the necessary columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("J2:J" & lastRow).NumberFormat = "0.00"
    Range("K1").Value = "Percentage Change"
    Range("K2:K" & lastRow).NumberFormat = "0.00%"
    Range("L1").Value = "Total Stock Volume"
    Range("J:L").Columns.AutoFit

    Range("O2").Value = "Greatest % Increase"
    Range("Q2").NumberFormat = "0.00%"
    Range("O3").Value = "Greatest % Decrease"
    Range("Q3").NumberFormat = "0.00%"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O:O").Columns.AutoFit

    'Assign starting values for the opening price of the first stock to be analyzed and first row to display its summary metrics
    startPrice = Range("C2").Value
    summaryRow = 2

    'Loop through the dataset starting from row 2 to the last row of the sheet's data set
    Dim i As Double
    For i = 2 To lastRow
    
        'If the current row is the last row of data for a certain stock (based on the ticker symbol)..
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            endPrice = Cells(i, 6).Value
   
            Ticker = Cells(i, 1).Value
            qChange = endPrice - startPrice
            pChange = qChange / startPrice
            totalVolume = totalVolume + Cells(i, 7).Value
       
        'Determine the color formatting for the percentage change metrics conditional on its value
                If qChange < 0 Then
                    colorIndex = 3
                ElseIf qChange > 0 Then
                    colorIndex = 4
                Else
                    colorIndex = 0
                End If
       
        'Print the summary metrics for the stock with conditional color formatting
            Cells(summaryRow, 9).Value = Ticker
            Cells(summaryRow, 10).Value = qChange
            Cells(summaryRow, 10).Interior.colorIndex = colorIndex
            Cells(summaryRow, 11).Value = pChange
            Cells(summaryRow, 12).Value = totalVolume
    
        'Track the stocks with the greatest % increase and greatest % decrease as the loop iterates
                If pChange > greatestIncrease Then
                    greatestIncrease = pChange
                    greatestIncTicker = Ticker
                ElseIf pChange < greatestDecrease Then
                    greatestDecrease = pChange
                    greatestDecTicker = Ticker
                End If
        
        'Track the stock with the greatest volume traded as the loop iterates
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolTicker = Ticker
                End If
        
        'Set the position for the next row of summary metrics
            summaryRow = summaryRow + 1
    
        'Reset the volume count and set the new starting price of the next stock to be analyzed
            totalVolume = 0
            startPrice = Cells(i + 1, 3).Value
    
        'If the current row is not the last row of data for a stock, update the total volume traded metric
        Else
            totalVolume = totalVolume + Cells(i, 7).Value

        End If
    
    Next i

    'Print the summary metrics for the top stocks based on % changes and volume traded
    Range("P2").Value = greatestIncTicker
    Range("Q2").Value = greatestIncrease
    Range("P3").Value = greatestDecTicker
    Range("Q3").Value = greatestDecrease
    Range("P4").Value = greatestVolTicker
    Range("Q4").Value = greatestVolume

    'Reset values of the "greatest" metrics for the next worksheet'
    greatestIncTicker = ""
    greatestIncrease = 0
    greatestDecTricker = ""
    greatestDecrease = 0
    greatestVolTicker = ""
    greatestVolume = 0

Next h

End Sub
