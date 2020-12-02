Attribute VB_Name = "Module1"
Sub MultiYearStock()

For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim Yearly_Change As Double
    Dim Yearly_Start As Double
    Dim Yearly_End As Double
    Dim Start_Row As Double
    Dim Percent_Change As Double
  
    
    
    Dim Ticker_Name As String
    
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    Dim Row_Count As Double
    Row_Count = 0
    
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    Dim x As Long
    
    For x = 2 To LastRow
    
        If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
        
        Ticker = ws.Cells(x, 1).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        
        Start_Row = x - Row_Count
        
        If ws.Cells(x, 6).Value Or ws.Range("C" & Start_Row).Value <> 0 Then
            
            Yearly_Open = ws.Cells(x, 3).Value
            Yearly_End = ws.Cells(x, 6).Value
            Yearly_Start = ws.Range("C" & Start_Row).Value
            Yearly_Change = Yearly_End - Yearly_Start
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            Percent_Change = Yearly_Change / Yearly_Start
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            End If
            
            
                Stock_Volume = Stock_Volume + ws.Cells(x, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                Row_Count = 0
                Stock_Volume = 0
                
            Else
            
                Row_Count = Row_Count + 1
                Stock_Volume = Stock_Volume + ws.Cells(x, 7).Value
            
            End If
            
                If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                                
                End If
                
    Next x
            

Next ws



End Sub
