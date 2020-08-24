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
    
    lastTicker = Cells(Rows.Count, 9).End(xlUp).Row
    
    For ticker_r = 2 To lastTicker
        Ticker = Cells(ticker_r, 9).Value
        YearlyChange = Cells(ticker_r, 10).Value
        PercentChange = Cells(ticker_r, 11).Value
        TotalVolume = Cells(ticker_r, 12).Value
        
        If (TotalVolume > GreatestTotalVolume) Then
            GreatestTotalVolume = TotalVolume
            VolumeTicker = Ticker
            Cells(4, 16).Value = Ticker
            Cells(4, 17).Value = GreatestTotalVolume
        End If
    Next ticker_r
    

    
    ' Auto fit columns
    Columns("O:Q").AutoFit
    
    
    
End Sub
