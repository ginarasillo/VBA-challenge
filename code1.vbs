Sub MultipleYearStockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        Dim i As Long
        
        Dim j As Long
       
        Dim TickerN As Long
      
        Dim LastRow As Long
       
        Dim LastRowB As Long
     
        Dim PercentChange As Double
       
        Dim GreatIncrease As Double
        
        Dim GreatDecrease As Double
       
        Dim GreatVolume As Double
       
        WorksheetName = ws.Name
        
        'column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        TickerN = 2
        
       
        j = 2
        
      
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
            'Loop
            For i = 2 To LastRow
            
               
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
               
                ws.Cells(TickerN, 9).Value = ws.Cells(i, 1).Value
                
                
                ws.Cells(TickerN, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                   
                    If ws.Cells(TickerN, 10).Value < 0 Then
                
                 
                    ws.Cells(TickerN, 10).Interior.ColorIndex = 3
                
                    Else
                
                   
                    ws.Cells(TickerN, 10).Interior.ColorIndex = 4
                
                    End If
                    
                   
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'formating
                    ws.Cells(TickerN, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerN, 11).Value = Format(0, "Percent")
                    
                    End If
              
                ws.Cells(TickerN, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
              
                TickerN = TickerN + 1
                
                
                j = i + 1
                
                End If
            
            Next i
            
   
        LastRowB = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
      
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop
            For i = 2 To LastRowB
            
                
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
               
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
           
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = GreatVolume
            
            Next i
     
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub