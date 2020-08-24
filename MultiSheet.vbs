Sub EachWSData()
    
    Dim TickerSymbol As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    Dim NewTicker As Boolean
    Dim Data_r As Integer  ' Which row to print data to
    
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

    Next ws

End Sub

