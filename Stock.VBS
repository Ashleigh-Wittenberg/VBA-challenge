Sub StockAnalysis()
    
    For Each ws In Worksheets

        Dim lastRow As Long
    
        Dim Ticker As String
        
        Dim OpeningPrice As Double
    
        Dim Difference As Double
        
        Dim DecimalChange As Double
    
        Dim YearlyStatsTable As Integer
        
        Dim Percent As String
        
        Dim MaxInc As Double
        
        Dim MaxDec As Double
        
        Dim MaxVol As Double
        
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        
        ws.Cells(3, 15).Value = "Greates % Decrease"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        
             YearlyStatsTable = 2
             
             TotalStock = 0
             
             MaxInc = 0
             
             MaxDec = 0
             
             MaxVol = 0
            
            OpeningPrice = ws.Cells(2, 3).Value
        

                For i = 2 To lastRow
                
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                        Ticker = ws.Cells(i, 1).Value
                        
                        ws.Range("I" & YearlyStatsTable).Value = Ticker
                
                        Difference = ws.Cells(i, 6).Value - OpeningPrice
                
                        ws.Cells(YearlyStatsTable, 10).Value = Difference
                        
                        If ws.Cells(YearlyStatsTable, 10).Value > 0 Then
                        
                            ws.Cells(YearlyStatsTable, 10).Interior.ColorIndex = 4
                            
                        Else
                        
                            ws.Cells(YearlyStatsTable, 10).Interior.ColorIndex = 3
                            
                        End If
                        
                        DecimalChange = Difference / OpeningPrice

                        ActualPercent = FormatPercent(DecimalChange, 2)
                        
                        ws.Cells(YearlyStatsTable, 11).Value = ActualPercent
                        
                        ws.Cells(YearlyStatsTable, 12).Value = TotalStock
                        
                        If ws.Cells(YearlyStatsTable, 11).Value > MaxInc Then
                        
                            MaxInc = ws.Cells(YearlyStatsTable, 11).Value
                            
                            MaxIncTicker = ws.Cells(YearlyStatsTable, 9).Value
                            
                        End If
                        
                        If ws.Cells(YearlyStatsTable, 11).Value < MaxDec Then
                        
                            MaxDec = ws.Cells(YearlyStatsTable, 11).Value
                            
                            MaxDecTicker = ws.Cells(YearlyStatsTable, 9).Value
                            
                        End If
                        
                        If ws.Cells(YearlyStatsTable, 12).Value > MaxVol Then
                        
                            MaxVol = ws.Cells(YearlyStatsTable, 12).Value
                            
                            MaxVolTicker = ws.Cells(YearlyStatsTable, 9).Value
                            
                        End If
                    
                        YearlyStatsTable = YearlyStatsTable + 1
                        
                        TotalStock = 0
                    
                        OpeningPrice = ws.Cells(i + 1, 3).Value
                    
                    End If
                
                Next i
                
                ws.Cells(2, "P").Value = MaxIncTicker
                
                ws.Cells(2, "Q").Value = FormatPercent(MaxInc, 2)
                
                ws.Cells(3, "P").Value = MaxDecTicker
                
                ws.Cells(3, "Q").Value = FormatPercent(MaxDec, 2)
                
                ws.Cells(4, "P").Value = MaxVolTicker
                
                ws.Cells(4, "Q").Value = MaxVol
                
        Next ws


End Sub