Sub GetTicker()
    Dim TickerSymbol As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    Dim NewTicker As Boolean
    Dim Data_r As Integer  ' Which row to print data to
    
    ' Declare opening variables for program to function
    NewTicker = True
    Data_r = 2
    
        ' Determine last row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Print headers for data collected
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        'Cells(1, 13).Value = "Year Begin Value"
        'Cells(1, 14).Value = "Year Close Value"
        

        '  Go through each row to collect data
        For r = 2 To LastRow
            If (NewTicker = True) Then
                TotalStock = 0
                NewTicker = False
                TickerSymbol = Cells(r, 1).Value
                YearOpen = Cells(r, 3).Value
            End If
            
            TotalStock = TotalStock + Cells(r, 7).Value
        
            If (Cells(r, 1).Value <> Cells(r + 1, 1).Value) Then
                NewTicker = True
                YearClose = Cells(r, 6).Value
                ' Calculate Yearly Change
                YearlyChange = YearClose - YearOpen
                ' Calculate Percent Change
                PercentChange = YearClose / YearOpen - 1
            End If
            
            ' Print data
            If (NewTicker = True) Then
                Cells(Data_r, 9).Value = TickerSymbol
                Cells(Data_r, 10).Value = YearlyChange
                Cells(Data_r, 11).Value = PercentChange
                Cells(Data_r, 11).NumberFormat = "0.00%" ' Make percentage have two decimal places
                'Cells(Data_r, 11).Style = "Percent"
                Cells(Data_r, 12).Value = TotalStock
                'Cells(Data_r, 13).Value = YearOpen
                'Cells(Data_r, 14).Value = YearClose
                
                ' Color cells for Yearly Change
                If (YearlyChange < 0) Then
                    Cells(Data_r, 10).Interior.ColorIndex = 3
                ElseIf (YearlyChange > 0) Then
                    Cells(Data_r, 10).Interior.ColorIndex = 4
                End If
                
                
                ' Set to next row forext ticker
                Data_r = Data_r + 1
            End If
        
        Next r
        
        ' Auto Fit Result columns
        Columns("I:L").AutoFit

End Sub

