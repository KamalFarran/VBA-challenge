Attribute VB_Name = "Module1"
Sub YearlySummary():
    
    Dim thisTicker As String
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    Dim Volume As Double
    
    Dim lastRow As Long
    Dim summaryRow As Integer
    
    Dim MinChange As Double
    Dim MaxChange As Double
    Dim MaxVolume As Double
    Dim MinChangeTicker As String
    Dim MaxChangeTicker As String
    Dim MaxVolumeTicker As String
    
    For Each ws In Worksheets
    
        'initialize summary headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'reset running volume and summary row position values
        Volume = 0
        summaryRow = 1
        
        'Get last row number
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Get the opening price and volume of the first row
        OpenPrice = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
        
            'Add the volume of the current row to the running volume
            Volume = Volume + ws.Cells(i, 7)
            
            'If there is a change of ticker value
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            
                'Move to new row in the summary section
                summaryRow = summaryRow + 1
                
                'Get the closing value and ticker name from the last row of the ticker
                ws.Cells(summaryRow, 9).Value = ws.Cells(i, 1).Value
                
                'Get the closing value to calculate year change
                ws.Cells(summaryRow, 10).Value = ws.Cells(i, 6).Value - OpenPrice
                YearlyChange = (ws.Cells(i, 6).Value - OpenPrice) / OpenPrice
                ws.Cells(summaryRow, 11).Value = YearlyChange
                
                'Change color according to year change
                If YearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Copy total volume from running volume
                ws.Cells(summaryRow, 12).Value = Volume
                
                'Reset the running volume and get the open price for the new ticker
                OpenPrice = ws.Cells(i + 1, 3).Value
                Volume = 0
                
            End If
            
            
        Next i
        
        'Format Cells
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("I:L").EntireColumn.AutoFit
        
        'Get PercentChange and Volume values from the first summary row
        MinChange = ws.Cells(2, 11)
        MaxChange = ws.Cells(2, 11)
        MaxVolume = ws.Cells(2, 12)
        MinChangeTicker = ws.Cells(2, 9)
        MaxChangeTicker = ws.Cells(2, 9)
        MaxVolumeTicker = ws.Cells(2, 9)
        
        'loop through all remaining summary rows
        For i = 3 To summaryRow
        
            'If percentage change is less than current minimum change, update minimum change data
            If ws.Cells(i, 11) < MinChange Then
                MinChange = ws.Cells(i, 11)
                MinChangeTicker = ws.Cells(i, 9)
            End If
            'If percentage change is more than current maximum change, update maximum change data
            If ws.Cells(i, 11) > MaxChange Then
                MaxChange = ws.Cells(i, 11)
                MaxChangeTicker = ws.Cells(i, 9)
            End If
            'If volume is greater than current maximum volume, update maximum volume data
            If ws.Cells(i, 12) > MaxVolume Then
                MaxVolume = ws.Cells(i, 12)
                MaxVolumeTicker = ws.Cells(i, 9)
            End If
            
        Next
        
        'Create Summary 2 headers
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        
        'Fill Summary 2 data
        ws.Cells(2, 16) = MaxChangeTicker
        ws.Cells(3, 16) = MinChangeTicker
        ws.Cells(4, 16) = MaxVolumeTicker
        ws.Cells(2, 17) = MaxChange
        ws.Cells(3, 17) = MinChange
        ws.Cells(4, 17) = MaxVolume
        
        
        'Format Summary2 Cells
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("O:Q").EntireColumn.AutoFit
        
    Next ws
    
    
End Sub
