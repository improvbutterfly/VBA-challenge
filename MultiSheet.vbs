Sub EachWSData()
    ' Declare variables needed for finding yearly change, percent change, and total stock volume
    Dim TickerSymbol As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    Dim NewTicker As Boolean
    Dim Data_r As Integer  ' Which row to print data to
    
    ' Declare variables needed for finding greatest % increase, greatest % decrease, and greatest total volume
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim IncreaseTicker As String
    Dim DecreaseTicker As String
    Dim VolumeTicker As String
    Dim Ticker As String
    Dim TotalVolume As Double
    
    
    For Each ws In Worksheets
        ' Declare opening variables for program to function
        NewTicker = True
        Data_r = 2
    
        ' Determine last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Print headers for data collected
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        '  Go through each row to collect data
        For r = 2 To LastRow
            If (NewTicker = True) Then
                TotalStock = 0
                NewTicker = False
                TickerSymbol = ws.Cells(r, 1).Value
                YearOpen = ws.Cells(r, 3).Value
            End If
            
            TotalStock = TotalStock + ws.Cells(r, 7).Value
        
            If (ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value) Then
                NewTicker = True
                YearClose = ws.Cells(r, 6).Value
                ' Calculate Yearly Change
                YearlyChange = YearClose - YearOpen
                
                ' In case data includes 0 for year open, need to test so program doesn't crash
                If (YearOpen <> 0) Then ' Test for zero value in year open
                ' Calculate Percent Change
                    PercentChange = YearClose / YearOpen - 1
                Else
                    PercentChange = 0
                End If
            End If
            
            ' Print data
            If (NewTicker = True) Then
                ws.Cells(Data_r, 9).Value = TickerSymbol
                ws.Cells(Data_r, 10).Value = YearlyChange
                ws.Cells(Data_r, 11).Value = PercentChange
                ws.Cells(Data_r, 11).NumberFormat = "0.00%" ' Make percentage have two decimal places
                ws.Cells(Data_r, 12).Value = TotalStock
                
                ' Color cells for Yearly Change
                If (YearlyChange < 0) Then
                    ws.Cells(Data_r, 10).Interior.ColorIndex = 3
                ElseIf (YearlyChange > 0) Then
                    ws.Cells(Data_r, 10).Interior.ColorIndex = 4
                End If
                
                ' Set to next row forext ticker
                Data_r = Data_r + 1
            End If
        
        Next r
        
        ' Auto Fit Result columns
        ws.Columns("I:L").AutoFit
        
        ' Move onto extra challenge to find the greatest % increase, greatest % decrease, and greatest total volume
        
        ' Print headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        ' Get the number of rows of the different tickers
        lastTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' For each ticker value
        For ticker_r = 2 To lastTicker
            Ticker = ws.Cells(ticker_r, 9).Value
            YearlyChange = ws.Cells(ticker_r, 10).Value
            PercentChange = ws.Cells(ticker_r, 11).Value
            TotalVolume = ws.Cells(ticker_r, 12).Value
        
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
        ws.Cells(2, 16).Value = IncreaseTicker
        ws.Cells(2, 17).Value = GreatestPercentIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%" ' Make percentage have two decimal places
    
        ' Display data for greatest percent decrease
        ws.Cells(3, 16).Value = DecreaseTicker
        ws.Cells(3, 17).Value = GreatestPercentDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%" ' Make percentage have two decimal places
    
        ' Display data for greatest total volume
        ws.Cells(4, 16).Value = VolumeTicker
        ws.Cells(4, 17).Value = GreatestTotalVolume
        
        ' Reset values for next worksheet
        GreatestPercentDecrease = 0
        GreatestPercentIncrease = 0
        GreatestTotalVolume = 0

    
        ' Auto fit columns
        ws.Columns("O:Q").AutoFit
    

    Next ws

End Sub
