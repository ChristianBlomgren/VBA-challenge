VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()

    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Integer
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of data in the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set initial values for summary table
        summaryTableRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Reset variables for each worksheet
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        
        ' Loop through all rows of data in the worksheet
        For i = 2 To lastRow
            ' Check if the current ticker symbol is different from the previous ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Set the opening price
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            ' Check if the current ticker symbol is the same as the next ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Set the closing price
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate the percent change
                If openingPrice <> 0 Then
                    percentChange = yearlyChange / openingPrice
                Else
                    percentChange = 0
                End If
                
                ' Add the yearly change, percent change, and total volume to the summary table
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 12).Value = totalVolume
                
                ' Format the percent change as a percentage
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                
                ' Color the yearly change cell based on positive or negative value
                If yearlyChange >= 0 Then
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                 ' Check if the current stock has the greatest increase
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                
                ' Check if the current stock has the greatest decrease
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                
                ' Check if the current stock has the greatest total volume
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Reset the total volume for the next ticker symbol
                totalVolume = 0
                
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
            End If
            
            ' Add the volume to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
        
         ' Print the stocks with the greatest increase, decrease, and volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
    Next ws
    
End Sub

