Attribute VB_Name = "Module1"
' ZiAng Zhang, 20181117
Private Type Pair
    Name As String
    Volume As LongLong
    Number As Double
End Type

Sub totalStockVolume()
    'Loop over all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Call totalStockVolumeSingleSheet
    Next
End Sub


Private Sub totalStockVolumeSingleSheet()
    'Function to compute total stock volume and stuff
    
    'Define line number to denote the line we are working at
    Dim LineNumber As Long
    Dim ResultLineNumber As Integer
    LineNumber = 1
    ResultLineNumber = 1
    'Store the name and volume of the stock we are working at
    Dim StockName As String
    Dim Volume As LongLong
    Dim TotalVolume As LongLong
    TotalVolume = 0
    Volume = 0
    'Define the opening and closing price
    Dim PriceOpening As Double
    Dim PriceClosing As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    PriceOpening = 0
    PriceClosing = 0
    YearlyChange = 0
    PercentChange = 0
    'Define the three special stocks that have the extreme values
    Dim GreatestPercentIncrease As Pair
    Dim GreatestPercentDecrease As Pair
    Dim GreatestTotalVolume As Pair
    GreatestPercentIncrease.Number = 0
    GreatestPercentDecrease.Number = 0
    GreatestTotalVolume.Volume = 0
    
    'Loop over each Row, as long as there is content
    While (IsEmpty(Cells(LineNumber, 1)) = False)
        'If it's different from the next line, either because we are at the first line, or that we have reached the end of a stock
        If (Cells(LineNumber + 1, 1).Value <> Cells(LineNumber, 1).Value) Then
            'If it's the title row
            If (LineNumber = 1) Then
                'Mostly do nothing, except for printing title row
                Cells(LineNumber, 11).Value = "<Ticker>"
                Cells(LineNumber, 12).Value = "Yearly Change"
                Cells(LineNumber, 13).Value = "Percent Change"
                Cells(LineNumber, 14).Value = "Total Volume"
                Cells(LineNumber, 18).Value = "<Ticker>"
                Cells(LineNumber, 19).Value = "Value"
            'End of a stock, Write output
            Else
                StockName = Cells(LineNumber, 1).Value
                Volume = Cells(LineNumber, 7).Value
                TotalVolume = TotalVolume + Volume
                'Get the closing price of this stock
                PriceClosing = Cells(LineNumber, 6).Value
                'Yearly Change
                YearlyChange = PriceClosing - PriceOpening
                If (PriceOpening = 0) Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / PriceOpening
                End If
                'Write result line
                ResultLineNumber = ResultLineNumber + 1
                Cells(ResultLineNumber, 11).Value = StockName
                Cells(ResultLineNumber, 12).Value = YearlyChange
                Cells(ResultLineNumber, 12).NumberFormat = "0.00"
                Cells(ResultLineNumber, 13).Value = PercentChange
                Cells(ResultLineNumber, 13).NumberFormat = "0.00%"
                'Choose color for the block
                If (YearlyChange >= 0) Then
                    Cells(ResultLineNumber, 12).Interior.Color = RGB(0, 250, 0)
                Else
                    Cells(ResultLineNumber, 12).Interior.Color = RGB(250, 0, 0)
                End If
                Cells(ResultLineNumber, 14).Value = TotalVolume
                'Since we just wrote output, time to check if this stock had the greatest changes
                If (PercentChange > GreatestPercentIncrease.Number) Then
                    GreatestPercentIncrease.Name = StockName
                    GreatestPercentIncrease.Number = PercentChange
                End If
                If (PercentChange < GreatestPercentDecrease.Number) Then
                    GreatestPercentDecrease.Name = StockName
                    GreatestPercentDecrease.Number = PercentChange
                End If
                If (TotalVolume > GreatestTotalVolume.Volume) Then
                    GreatestTotalVolume.Name = StockName
                    GreatestTotalVolume.Volume = TotalVolume
                End If
                'Clear out total volume
                TotalVolume = 0
                Volume = 0
            End If
            'Record the opening price of the following stock
            PriceOpening = Cells(LineNumber + 1, 3).Value
            PriceClosing = 0
        'If it's a row within one stock
        Else
            StockName = Cells(LineNumber, 1).Value
            Volume = Cells(LineNumber, 7).Value
            TotalVolume = TotalVolume + Volume
        End If
        LineNumber = LineNumber + 1
    Wend
    
    'Write out the three most outstanding stocks
    Cells(2, 17).Value = "Greatest Percent Increase"
    Cells(2, 18).Value = GreatestPercentIncrease.Name
    Cells(2, 19).Value = GreatestPercentIncrease.Number
    Cells(2, 19).NumberFormat = "0.00%"
    Cells(2, 19).Interior.Color = RGB(0, 250, 0)
    
    Cells(3, 17).Value = "Greatest Percent Decrease"
    Cells(3, 18).Value = GreatestPercentDecrease.Name
    Cells(3, 19).Value = GreatestPercentDecrease.Number
    Cells(3, 19).NumberFormat = "0.00%"
    Cells(3, 19).Interior.Color = RGB(250, 0, 0)
    
    Cells(4, 17).Value = "Greatest Total Volume"
    Cells(4, 18).Value = GreatestTotalVolume.Name
    Cells(4, 19).Value = GreatestTotalVolume.Volume
End Sub

