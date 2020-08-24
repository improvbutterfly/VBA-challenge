Sub GetTicker()
    Dim TickerSymbol As String
    Dim YearOpen As Long
    Dim YearClose As Long
    Dim YearlyChange As Long
    Dim PercentChange As Long
    Dim TotalStock As Double
    Dim NewTicker As Boolean
    Dim Data_r As Integer  ' Which row to print data to
    
    ' Declare opening variables for program to function
    NewTicker = True
    TotalStock = 0
    Data_r = 2
    
        ' Determine last row
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Print headers for data collected
        Cells(1, 9).Value = "Ticker"
        'Cells(1, 10).Value = "Yearly Change"
        'Cells(1, 11).Value = "Percent Change"
        Cells(1, 10).Value = "Year Begin Value"
        Cells(1, 11).Value = "Year Close Value"
        Cells(1, 12).Value = "Total Stock Volume"
        

        '  Go through each row to collect data
        For r = 2 To LastRow
            If (NewTicker = True) Then
                NewTicker = False
                TickerSymbol = Cells(r, 1).Value
                YearOpen = Cells(r, 3).Value
            End If
        
            If (Cells(r, 1).Value <> Cells(r + 1, 1).Value) Then
                NewTicker = True
                YearClose = Cells(r, 6).Value
            End If
            
            ' Print data
            If (NewTicker = True) Then
                Cells(Data_r, 9).Value = TickerSymbol
                Cells(Data_r, 10).Value = YearOpen
                Cells(Data_r, 11).Value = YearClose
                Data_r = Data_r + 1
            End If
        
        Next r

End Sub


