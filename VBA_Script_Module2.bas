Attribute VB_Name = "Module1"
Sub TickerFormula()
    
    For Each ws In Worksheets
    
    Dim WorksheetName As String
    
    
    
    WorksheetName = ws.Name
    
    
    Dim Ticker_Name As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume As Double
    
    Volume = 0
    
    Dim open_price As Double
    
    open_price = ws.Cells(2, 3).Value
    
    Dim year_high As Double
    Dim year_low As Double
    Dim close_price As Double
    
    
    
    

    
    
    
    Dim SummaryTicker As Integer
    SummaryTicker = 2
    
    
    
    Dim LR As Long
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker Name"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To LR
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'This prints Ticker Name in New Column
            Ticker_Name = ws.Cells(i, 1).Value
            
            'This Adds Total Stock Vol
            Volume = Volume + ws.Cells(i, 7).Value
            
            ws.Range("I" & SummaryTicker).Value = Ticker_Name
            
            ws.Range("L" & SummaryTicker).Value = Volume
            
            close_price = ws.Cells(i, 6).Value
            
            Yearly_Change = (close_price - open_price)
            
            ws.Range("J" & SummaryTicker).Value = Yearly_Change
                
                If (open_price = 0) Then
                
                    Percent_Change = 0
                    
                Else
                    
                    Percent_Change = Yearly_Change / open_price
                    
                End If
                
            
            
            
            
            ws.Range("K" & SummaryTicker).Value = Percent_Change
            ws.Range("K" & SummaryTicker).NumberFormat = "0.00%"
                
             
            
            SummaryTicker = SummaryTicker + 1
            
            Volume = 0
            
            open_price = ws.Cells(i + 1, 3)
            
        Else
        
            Volume = Volume + ws.Cells(i, 7).Value
            
          
            
        
        End If
        
    
    Next i
    
    
    'Must Specify Column 9 in Last Row Formula to be able to determine last row in summary table and not in datasheet
    
    lastrow_SummaryTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
    
    For i = 2 To lastrow_SummaryTicker
            If ws.Range("J" & SummaryTicker).Value > 0 Then
                ws.Range("J" & SummaryTicker).Interior.ColorIndex = 4
            Else
                ws.Range("J" & SummaryTicker).Interior.ColorIndex = 3
            End If
    Next i
    
    
    
    lastrow_SummaryTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Greatest_Volume = ws.Cells(2, 12).Value
    Gr_Increase = ws.Cells(2, 11).Value
    Gr_Decrease = ws.Cells(2, 11).Value
    
    For i = 2 To lastrow_SummaryTicker
        If ws.Cells(i, 12).Value > Greatest_Volume Then
        
        Greatest_Volume = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        Greatest_Volume = Greatest_Volume
        
        End If
        
        
        If ws.Cells(i, 11).Value > Gr_Increase Then
        
        Gr_Increase = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        Gr_Increase = Gr_Increase
        
        End If
    
    
        If ws.Cells(i, 11).Value < Gr_Decrease Then
        
        Gr_Decrease = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        Gr_Decrease = Gr_Decrease
        
        End If
        
    Next i
    
        
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = Format(Greatest_Volume, "Scientific")
    
    
    
    
    
    Next ws















End Sub
