Sub Stocks()

    Dim ws As Worksheet
    Dim Last_Row As Long
    Dim ticker_index As String
    Dim year_open_index As Integer
    Dim year_close_index As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim percent_change As Integer
    Dim total_volume As Long
    Dim i As Long
    Dim year_change As Double
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_total_volume As Long

    For Each ws In Worksheets

        'Set indexes
        ticker_index = 2
        year_open_index = 3
        year_close_index = 6
        total_volume = 0
        greatest_percent_increase = 13
        greatest_percent_decrease = 14
        greatest_total_volume = 15
            
        'Define Loop End Range
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        'Loop through spreadsheets, rows and columns
    For i = 2 To LastRow
            'Define Variable Calculations
             year_open = ws.Cells(i, year_open_index).Value
             year_close = ws.Cells(i, year_close_index).Value
             year_change = (year_close - year_open)
             percent_change = (year_change / year_open) * 100
             total_volume = total_volume + ws.Cells(i, 7).Value
                    
               'check to see when each row has different ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               'display variable results in respective summary columns
                    ws.Cells(ticker_index, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(ticker_index, 10).Value = year_change
                    ws.Cells(ticker_index, 11).Value = percent_change
                    ws.Cells(ticker_index, 12).Value = total_volume
                   
                    If ws.Cells(i, 10).Value > ws.Cells(i + 1, 10) Then
                        ws.Cells(i, 10).Value = greatest_percent_increase
                    End If
                    If ws.Cells(i, 10).Value < ws.Cells(i + 1, 10) Then
                        ws.Cells(i, 10).Value = greatest_percent_decrease
                    End If
                    If ws.Cells(i, 12).Value > ws.Cells(i + 1, 10) Then
                        ws.Cells(i, 12).Value = greatest_total_volume
                    End If
                                      
                'Assign color filter
                If percent_change > 0 Then
                    ws.Columns(10).Interior.ColorIndex = 4
                Else
                    ws.Columns(10).Interior.ColorIndex = 3
    
                If year_change > 0 Then
                    ws.Columns(11).Interior.ColorIndex = 4
                Else
                    ws.Columns(11).Interior.ColorIndex = 3

                'Reset loop for next ticker
                    ticker_index = ticker_index + 1
                    total_volume = 0
                    
        End If
     Next ws
End Sub