Attribute VB_Name = "Module1"
Sub stocks():
Dim row As Integer
Dim ticker As String
Dim TotalStockVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim QuarterlyChange As Double
Dim PercentChange As Double


'Loop through all sheets
For Each ws In Worksheets

    'find the last row in each of the sheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'Add the column headers for I1, J1, K1, L1
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'start the row at the second row
    row = 2
    row2 = 2
    TotalStockVolume = 0
    
    'start a for loop to go through all the tickers information
    For i = 2 To LastRow
        
        'Collects ticker and opening price at the beginning of the quarter
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
            
            'collects the ticker value
            ticker = ws.Cells(i, 1).Value
            
            'collects the opening price at the beginning of the quarter
            OpenPrice = ws.Cells(i, 3).Value
            
            'outputs ticker into column I
            ws.Range("I" & row).Value = ticker
            
            'moves the row down
            row = row + 1
        
        End If
        
        'Collects closing price at the end of the quarter
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            ClosePrice = ws.Cells(i, 6).Value
            
            'calculate quarterly change based on beginning and end of quarter\
            QuarterlyChange = ClosePrice - OpenPrice
            
            'calculates total stock volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'calculates percent change
            PercentChange = QuarterlyChange / OpenPrice
            
            'displays quarterly change
            ws.Range("J" & row2).Value = QuarterlyChange
            
            'displays percent change
            ws.Range("K" & row2).Value = PercentChange
            
            'displays total stock volume
            ws.Range("L" & row2).Value = TotalStockVolume
            
            
            'resets total stock volume
            TotalStockVolume = 0
            
            'if positive turn the box green
            If ws.Cells(row2, 10).Value > 0 Then
                ws.Cells(row2, 10).Interior.ColorIndex = 4
                
            'if negative turn the box red
            ElseIf ws.Cells(row2, 10).Value < 0 Then
                ws.Cells(row2, 10).Interior.ColorIndex = 3
            
            'if it is zero do nothing
            Else
            End If
            
            'move the row by 1
            row2 = row2 + 1
            
        'If they are the same add the total stock volume
        Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

        End If
        
        
    Next i
    'gets greatest percent increase
    Dim MaxPercentIncrease As Double
    Dim TickerMaxPercentIncrease As String
    
    'gets greatest percent decrease
    Dim MaxPercentDecrease As Double
    Dim TickerMaxPercentDecrease As String
    
    'gets greatest total volume
    Dim MaxTotalVolume As Double
    Dim TickerMaxTotalVolume As String
    
    MaxPercentIncrease = 0
    MaxPercentDecrease = 0
    MaxTotalVolume = 0
    
    'loop through to find the max percent inc, dec, and total
    For i = 2 To LastRow - 1
        
        If ws.Cells(i, 11).Value > MaxPercentIncrease Then
            MaxPercentIncrease = ws.Cells(i, 11).Value
            TickerMaxPercentIncrease = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = MaxPercentIncrease
            ws.Cells(2, 16).Value = TickerMaxPercentIncrease
        End If
        
        If ws.Cells(i, 11).Value < MaxPercentDecrease Then
            MaxPercentDecrease = ws.Cells(i, 11).Value
            TickerMaxPercentDecrease = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = MaxPercentDecrease
            ws.Cells(3, 16).Value = TickerMaxPercentDecrease
        End If
        
        If ws.Cells(i, 12).Value > MaxTotalVolume Then
            MaxTotalVolume = ws.Cells(i, 12).Value
            TickerMaxTotalVolume = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = MaxTotalVolume
            ws.Cells(4, 16).Value = TickerMaxTotalVolume
        End If
        
    Next i
    
    ' format to make percent change an actual percent
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    'add row headers
    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % decrease"
    ws.Cells(4, 15) = "Greatest total volume"
    
    'add column headers
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
    ' format to make the columns autofit and look nicer
    ws.Columns("A:R").AutoFit
    ws.Cells(4, 17).NumberFormat = "0"



Next ws




End Sub
