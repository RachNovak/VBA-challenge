Attribute VB_Name = "Module1"
    Sub stocks()

For Each ws In Worksheets
    Dim Worksheet_Name As String
    Worksheet_Name = ws.Name
    
    ' Set an initial variable for holding the stock name
    Dim Stock_Name As String
    
    ' Set an initial variable for holding the total stock total
    Dim Stock_Total As Double
    Stock_Total = 0
    
    ' Keep track of the location for each stockname in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim lrow As Long
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
       
    'Set Table Headers:
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
       
    ' Loop through all stock data
    Dim i As Long
    For i = 2 To lrow
    
        ' Check if we are still within the same stock name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set the Stock name
        Stock_Name = ws.Cells(i, 1).Value
        
        ' Add to the Stock Total
        Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
        ' Print the Stock name in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Stock_Name
        
        ' Print the Stock Total to the Summary Table
       ws.Range("L" & Summary_Table_Row).Value = Stock_Total
          
        ' Set the Yearly change
        yopen = ws.Cells(2, 3).Value
        yclose = ws.Cells(i, 6).Value
        yrchange = yclose - yopen
            
        'Print the yearly changes in the Summary Table with correct formatting
        ws.Range("J" & Summary_Table_Row).Value = yrchange
            If ws.Range("J" & Summary_Table_Row).Value <= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(238, 75, 43)
            End If
            
            If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
            End If
            
            
        ' Set the percent changea
        popen = Cells(i, 3).Value
        pclose = Cells(i, 6).Value
        pcchange = ((pclose - popen) / popen)
            
            If ws.Cells(i, 3).Value <> 0 Then
            
            'Print the percent change in the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Format(pcchange, "Percent")

            End If
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Brand Total
            Stock_Total = 0
        
    ' If the cell immediately following a row is the same ticker name...
    Else
    
          ' Add to the Brand Total
          Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    
        End If
        
    Next i
    
    'Set Table Headers for the Greatest:
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "GreatestIncrease"
    ws.Cells(3, 14).Value = "GreatestDecerease"
    ws.Cells(4, 14).Value = "GreatestTotalVolume"

    
    Dim Table2 As Integer
    Table2 = 2
    
    Dim lrowmaxyc As Long
    lrowmaxyc = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    GreatestDecerease = ws.Cells(2, 11).Value
    GreatestIncrease = ws.Cells(2, 11).Value
    GreatestTotalVolume = ws.Cells(2, 12).Value

    Dim r As Long
    For r = 2 To lrowmaxyc
        
        'Greatest %  Increase
        If ws.Cells(r, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(r, 11).Value
                ws.Cells(2, 15).Value = ws.Cells(r, 9).Value
                
                Else
                GreatestIncrease = GreatestIncrease
                
                End If
                'print in correct format
                ws.Cells(2, 16).Value = Format(GreatestIncrease, "Percent")
        
        'Greatest % Decerease
         If ws.Cells(r, 11).Value < GreatestDecerease Then
                GreatestDecerease = ws.Cells(r, 11).Value
                ws.Cells(3, 15).Value = ws.Cells(r, 9).Value
                
                Else
                GreatestDecerease = GreatestDecerease
                
                End If
                'print in correct format
                ws.Cells(3, 16).Value = Format(GreatestDecerease, "Percent")
        
        'Greatest total volume
        If ws.Cells(r, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(r, 12).Value
                ws.Cells(4, 15).Value = ws.Cells(r, 9).Value
                
                Else
                GreatestTotalVolume = GreatestTotalVolume
                
                End If
                'print in table
                ws.Cells(4, 16).Value = (GreatestTotalVolume)
        
    Next r
           
    Next ws
    End Sub
    
    

