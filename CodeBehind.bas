Attribute VB_Name = "Module1"
Sub Stock_Data()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
  
        ' Add the word Ticker to Column "K" in the Header of each worksheet
        ws.Cells(1, 11).Value = "Ticker"

        ' Add the column label Total Stock Volume to Column "L" in the Header of each worksheet
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Add the column label Yearly Change to Column "M" in the Header of each worksheet
        ws.Cells(1, 13).Value = "Yearly Change"
        
        ' Add the column label Percent Change to Column "N" in the Header of each worksheet
        ws.Cells(1, 14).Value = "Percent Change"
        
        ' Add the ROW label Greatest Percent Increase to Column "P" in the Header of each worksheet
        ws.Cells(2, 16).Value = "Greatest Percent Increase"
        
        ' Add the ROW label Greatest Percent Decrease to Column "P" in the Header of each worksheet
        ws.Cells(3, 16).Value = "Greatest Percent Decrease"
        
        ' Add the ROW label Greatest Total Volume to Column "P" in the Header of each worksheet
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        ' Add the word Ticker to Column "Q" in the Header of each worksheet
        ws.Cells(1, 17).Value = "Ticker"
        
        ' Add the word Ticker to Column "R" in the Header of each worksheet
        ws.Cells(1, 18).Value = "Value"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set an initial variable for holding the ticker symbol
         Dim Ticker_Symbol As String
    ' ----------------------------------------
    ' MODERATE PART OF THE ASSIGNMENT
    ' ----------------------------------------
        ' Set an initial variable for holding the total amount of each stock per year
        Dim Stock_Volume_Total As Double
        Stock_Volume_Total = 0

        ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Set a variable to account for the Yearly change from what the stock opened the year at to what the closing price was.
        
        Dim YearlyChange As Double
        Dim LowestDate As Double
        Dim HighestDate As Double
        Dim OpenPrice As Double
        Dim ClosedPrice As Double
        Dim PercentChange As String
        Dim color As Integer
        
        
        LowestDate = 77777777
        HighestDate = 0
        color = 0
        
         ' Loop through each row to find the lowest date
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Calculate the yearlychange
            YearlyChange = ClosedPrice - OpenPrice
    
            'Set conditional formatting that will highlight positive change in green and negative change in red
            If YearlyChange >= 0 Then

                ' Set the Color of the cell to green
                  color = 4
            Else
                ' Set the Color of the cell to red
                  color = 3
            End If
            
               
            
            
            'The percent change from the what it opened the year at to what it closed.
            If OpenPrice = 0 Then
                PercentChange = Str((YearlyChange * 100)) + "%"
            Else
                PercentChange = Str(((YearlyChange / OpenPrice) * 100)) + "%"
            End If
            
            ' Print the YearlyChange in column M
            ws.Range("M" & Summary_Table_Row).Value = YearlyChange
            ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = color
            
            ws.Range("N" & Summary_Table_Row).Value = PercentChange
            
            LowestDate = 77777777
            HighestDate = 0
            YearlyChange = 0
            
        End If
    
     

        ' Loop through all stock entry items
           ' For i = 2 To LastRow

            ' Check if we are still grabbing the total volume amount as it relates to the ticker symbol
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Set the ticker symbol
                    Ticker_Symbol = ws.Cells(i, 1).Value
        
                    ' Add to the Total Stock Volume
                    Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value
        
                    ' Print the ticker symbole in the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Ticker_Symbol
        
                    ' Print the total stock volume Amount to the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
        
                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    ' Reset the Total Stock Volume
                    Stock_Volume_Total = 0
    
                ' If the cell immediately following a row is the same then proceed to the alternative option.
                Else
    
                ' Add to the Brand Total
                    Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value
    
                End If
                
        ' Compare dates to find lowest and highest date...
        If ws.Cells(i, 2).Value < LowestDate Then
        
            'Set LowestDate to the lowest value
             LowestDate = ws.Cells(i, 2).Value
             
             'Grab the OpenPrice
             OpenPrice = ws.Cells(i, 3).Value
        End If
        
         ' Compare dates to find highest
        If ws.Cells(i, 2).Value > HighestDate Then
        
            'Set LowestDate to the lowest value
             HighestDate = ws.Cells(i, 2).Value
             
             'Grab the ClosedPrice
             ClosedPrice = ws.Cells(i, 6).Value
        End If
            
        Next i
        
        ' ----------------------------------------
        ' HARD PART OF THE ASSIGNMENT
        ' ----------------------------------------
            
        'Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
       Dim GreatestPercentIncrease As Double
       Dim GreatestPercentDecrease As Double
       Dim GreatestTotalVolume As Double
       Dim gpiTicker As String
       Dim gpdTicker As String
       Dim gtvTicker As String
       
       GreatestPercentIncrease = -99999
       GreatestPercentDecrease = 99999
       GreatestTotalVolume = 0
       
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 13).End(xlUp).Row
    
        'For Loop
        For j = 2 To LastRow
            If ws.Cells(j, 13).Value > GreatestPercentIncrease Then
                GreatestPercentIncrease = ws.Cells(j, 13).Value
                gpiTicker = ws.Cells(j, 11).Value
            End If
            If ws.Cells(j, 13).Value < GreatestPercentDecrease Then
                GreatestPercentDecrease = ws.Cells(j, 13).Value
                gpdTicker = ws.Cells(j, 11).Value
            End If
            If ws.Cells(j, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = ws.Cells(j, 12).Value
                gtvTicker = ws.Cells(j, 11).Value
            End If
        Next j
        ws.Cells(2, 17).Value = gpiTicker
        ws.Cells(3, 17).Value = gpdTicker
        ws.Cells(4, 17).Value = gtvTicker
        ws.Cells(2, 18).Value = GreatestPercentIncrease
        ws.Cells(3, 18).Value = GreatestPercentDecrease
        ws.Cells(4, 18).Value = GreatestTotalVolume
        
    Next ws

End Sub















