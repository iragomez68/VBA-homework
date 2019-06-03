Sub CalculateStockKpi()
    Dim workhseetName As String
    'Loop thru all the sheets in the Workbook to calculate ticker's total
    'volume and percentChange
    For Each ws In ThisWorkbook.Worksheets
        worksheetName = ws.Name
        If worksheetName <> "" Then
            SetStockKpi (worksheetName)
        End If
    Next

End Sub

Sub SetStockKpi(ByVal sheetName As String)

Dim targetWorksheet As Worksheet
Dim tickerName As String
Dim tickerGreatestVolume As String
Dim tickerGreatestIncrease As String
Dim tickerGreatestDecrease As String

Dim tickerRow As Long
Dim resultRow As Long

Dim startDate As Long
Dim endDate As Long

Dim openingPrice As Double
Dim closingPrice As Double
Dim tickerVolume As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double

greatestIncrease = 0
greatestDecrease = 0
greatestVolume = 0

Set targetWorksheet = ThisWorkbook.Worksheets(sheetName)
tickerRow = 2
tickerVolume = 0
resultRow = 1
startDate = 99991231
endDate = 0
openingPrice = 0
closingPrice = 0

'Set the results header
targetWorksheet.Range("J" & resultRow).Value = "Ticker"
targetWorksheet.Range("K" & resultRow).Value = "Yearly Change"
targetWorksheet.Range("L" & resultRow).Value = "Percent Change"
targetWorksheet.Range("M" & resultRow).Value = "Total Stock Volume"

Do
    tickerName = targetWorksheet.Range("A" & tickerRow).Value
    tickerVolume = tickerVolume + targetWorksheet.Range("G" & tickerRow).Value
    
    'In case the data is not sorted, always keep track of the
    'earliest date and use the opening price in that row as the
    'new opening price
    
    If startDate > targetWorksheet.Range("B" & tickerRow).Value Then
         startDate = targetWorksheet.Range("B" & tickerRow).Value
         openingPrice = targetWorksheet.Range("C" & tickerRow).Value
    End If
        
    'In case the data is not sorted, always keep track of the
    'latest date and use the closing price in that row as the
    'new closing price
    
    If endDate < targetWorksheet.Range("B" & tickerRow).Value Then
         endDate = targetWorksheet.Range("B" & tickerRow).Value
         closingPrice = targetWorksheet.Range("F" & tickerRow).Value
    End If
        
    tickerRow = tickerRow + 1
    
    'Once the ticker changes then write the summary result
    'for the ticker
    If tickerName <> targetWorksheet.Range("A" & tickerRow).Value Then
        resultRow = resultRow + 1
        
        'calculate yearly change
        yearlyChange = closingPrice - openingPrice
        
        'Ensure there is no divide by 0
        If yearlyChange = 0 And openingPrice = 0 Then
            percentChange = 0
        ElseIf openingPrice = 0 Then
            percentChange = 1#
        Else
            percentChange = yearlyChange / openingPrice
        End If
        
        targetWorksheet.Range("J" & resultRow).Value = tickerName
        targetWorksheet.Range("K" & resultRow).Value = yearlyChange
        
        'set the cell color to green for positive or zero and red for negative
        If yearlyChange >= 0 Then
            targetWorksheet.Range("K" & resultRow).Interior.ColorIndex = 4
        Else
            targetWorksheet.Range("K" & resultRow).Interior.ColorIndex = 3
        End If
        
        targetWorksheet.Range("L" & resultRow).Value = percentChange
        targetWorksheet.Range("L" & resultRow).NumberFormat = "0.00%"
        
        targetWorksheet.Range("M" & resultRow).Value = tickerVolume
        
        'Bonus: keep track of the totalVolumen and percentChange across
        'all tickers.
        If tickerVolume > greatestVolume Then
            greatestVolume = tickerVolume
            tickerGreatestVolume = tickerName
        End If
        
        If percentChange > greatestIncrease Then
            greatestIncrease = percentChange
            tickerGreatestIncrease = tickerName
        End If
        
        If percentChange < greatestDecrease Then
            greatestDecrease = percentChange
            tickerGreatestDecrease = tickerName
        End If
        
        'reset variables to start summarazing next ticker
        tickerVolume = 0
        yearlyChange = 0
        percentChange = 0
        startDate = 99991231
        endDate = 0
        tickerName = targetWorksheet.Range("A" & tickerRow).Value
    End If
    
Loop Until tickerName = ""

targetWorksheet.Range("P" & 2).Value = "Greatest % increase"
targetWorksheet.Range("P" & 3).Value = "Greatest % Decrease"
targetWorksheet.Range("P" & 4).Value = "Greatest total volume"
targetWorksheet.Range("Q" & 1).Value = "Ticker"
targetWorksheet.Range("R" & 1).Value = "Value"


targetWorksheet.Range("Q" & 2).Value = tickerGreatestIncrease
targetWorksheet.Range("Q" & 3).Value = tickerGreatestDecrease
targetWorksheet.Range("Q" & 4).Value = tickerGreatestVolume


targetWorksheet.Range("R" & 2).Value = greatestIncrease
targetWorksheet.Range("R" & 2).NumberFormat = "0.00%"
targetWorksheet.Range("R" & 3).Value = greatestDecrease
targetWorksheet.Range("R" & 3).NumberFormat = "0.00%"
targetWorksheet.Range("R" & 4).Value = greatestVolume



End Sub


