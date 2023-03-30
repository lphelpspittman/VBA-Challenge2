Sub AnalyzeStocks():

    Dim row As Long
    Dim RowCount As Long
    Dim nextRow As Long
    Dim totalStockVolume As Double
    Dim openingValue As Double
    Dim closingValue As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim ws As Worksheet
    Dim rownumber As Integer
    Dim greatestPercentIncreaseValue As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseValue As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalStockVolumeTicker As String
    Dim greatestTotalStockVolumeValue As Double
    
    
    'Set ws = ActiveSheet
    For Each ws In Sheets
    
    totalStockVolume = 0
    openingValue = 0
    closingValue = 0
    YearlyChange = 0
    PercentChange = 0
    greatestTotalStockVolumeTicker = "<none>"
    greatestTotalStockVolumeValue = 0
    greatestPercentIncreaseValue = 0
    greatestPercentIncreaseTicker = "<none>"
    greatestPercentDecreaseValue = 0
    greatestPercentDecreaseTicker = "<none>"
    nextRow = 2
    
    
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
   
    'Get the row number of the last row with data
    RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    'Loop through all Ticker symbols...
    For row = 2 To RowCount
    
        'Check if we are beggining a ticker symbol block...
        If ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
            totalStockVolume = 0
        
        openingValue = ws.Cells(row, 3).Value
        
        End If
        
        
        'For each row add the total stock volume...
        totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value
            
            
            'save the closing price
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            closingValue = ws.Cells(row, 6).Value
            PercentChange = (closingValue - openingValue) / openingValue
            ws.Cells(nextRow, 9).Value = ws.Cells(row, 1).Value
            ws.Cells(nextRow, 10).Value = closingValue - openingValue
            ws.Cells(nextRow, 11).Value = PercentChange
            ws.Cells(nextRow, 12).Value = totalStockVolume
            
        If closingValue > openingValue Then
            ws.Cells(nextRow, 10).Interior.Color = vbGreen
        ElseIf closingValue < openingValue Then
            ws.Cells(nextRow, 10).Interior.Color = vbRed
        
        ws.Range("$J:$J").NumberFormat = "$#,##0.00"
        ws.Range("$K:$K").NumberFormat = "0.00%"
        ws.Range("Q2", "Q3").NumberFormat = "0.00%"
        
        End If
        
         
        If PercentChange > greatestPercentIncreaseValue Then
            greatestPercentIncreaseTicker = ws.Cells(row, 1).Value
            greatestPercentIncreaseValue = PercentChange
            
            
        End If
        
        If PercentChange < greatestPercentDecreaseValue Then
            greatestPercentDecreaseTicker = ws.Cells(row, 1).Value
            greatestPercentDecreaseValue = PercentChange
            
        End If
        
        If totalStockVolume > greatestTotalStockVolumeValue Then
            greatestTotalStockVolumeTicker = ws.Cells(row, 1).Value
            greatestTotalStockVolumeValue = totalStockVolume
            
       
        
        End If
            
             nextRow = nextRow + 1
             
        End If
        Next row
        
        ws.Range("P4").Value = greatestTotalStockVolumeTicker
        ws.Range("Q4").Value = greatestTotalStockVolumeValue
        ws.Range("P2").Value = greatestPercentIncreaseTicker
        ws.Range("Q2").Value = greatestPercentIncreaseValue
        ws.Range("P3") = greatestPercentDecreaseTicker
        ws.Range("Q3") = greatestPercentDecreaseValue
            
        
 Next ws
        
End Sub




