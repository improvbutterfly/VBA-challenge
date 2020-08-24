Sub CalculateGreatest()
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim IncreaseTicker As String
    Dim DecreaseTicker As String
    Dim VolumeTicker As String
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
        
    
    ' Print headers
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    ' Get the number of rows of the different tickers
    lastTicker = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' For each ticker value
    For ticker_r = 2 To lastTicker
        Ticker = Cells(ticker_r, 9).Value
        YearlyChange = Cells(ticker_r, 10).Value
        PercentChange = Cells(ticker_r, 11).Value
        TotalVolume = Cells(ticker_r, 12).Value
        
        ' Compare total volume to greatest volume
        If (TotalVolume > GreatestTotalVolume) Then
            GreatestTotalVolume = TotalVolume
            VolumeTicker = Ticker
        End If
        
        ' Compare percent increase to greatest increase
        If (PercentChange > GreatestPercentIncrease) Then
            GreatestPercentIncrease = PercentChange
            IncreaseTicker = Ticker
        ' Compae percent decrease to greatest decrease
        ElseIf (PercentChange < GreatestPercentDecrease) Then
            GreatestPercentDecrease = PercentChange
            DecreaseTicker = Ticker
        End If
    Next ticker_r
    
    ' Display data for greatest percent increase
    Cells(2, 16).Value = IncreaseTicker
    Cells(2, 17).Value = GreatestPercentIncrease
    Cells(2, 17).NumberFormat = "0.00%" ' Make percentage have two decimal places
    
    ' Display data for greatest percent decrease
    Cells(3, 16).Value = DecreaseTicker
    Cells(3, 17).Value = GreatestPercentDecrease
    Cells(3, 17).NumberFormat = "0.00%" ' Make percentage have two decimal places
    
    ' Display data for greatest total volume
    Cells(4, 16).Value = VolumeTicker
    Cells(4, 17).Value = GreatestTotalVolume

    
    ' Auto fit columns
    Columns("O:Q").AutoFit
    
    
    
End Sub
