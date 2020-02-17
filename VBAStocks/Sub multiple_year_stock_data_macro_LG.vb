Sub multiple_year_stock_data_2014()

'Loop though all the sheets
For Each Worksheet In Worksheets

    'Name the table columns
     Worksheet.Cells(1, 9).Value = "Ticker"
     Worksheet.Cells(1, 10).Value = "Yearly Change"
     Worksheet.Cells(1, 11).Value = "Percent Change"
     Worksheet.Cells(1, 12).Value = "Total Stock Volume'"
     
    'Set an initial variable to hold the ticker
    Dim Ticker As String
    
    'Set an initial variable to hold the yearly change
    Dim Yearly_Change As Double
    
    'Set a variable for opening price
    Dim Opening_Price As Double
    
    'Set a variable for closing price
    Dim Closing_Price As Double
    
    'Set an initial variable to hold the percentage change
    Dim Percentage_Change As Double
    
    'Set an initial variable to hold the total volume
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    
        'Create variable for last row
        Lastrow = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        
        'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        
        'Loop though all volumes for each worksheet
        For i = 2 To Lastrow
        
            
            'Add to the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + Worksheet.Cells(i, 7).Value
        
                If Worksheet.Cells(i, 1).Value <> Worksheet.Cells(i - 1, 1).Value Then
                    'Get the opening price
                    Opening_Price = Worksheet.Cells(i, 3).Value
                End If
                If Worksheet.Cells(i, 1).Value <> Worksheet.Cells(i + 1, 1).Value Then
                    'Set the ticker type and get the closing price
                    Ticker = Worksheet.Cells(i, 1).Value
                    Closing_Price = Worksheet.Cells(i, 6).Value
                    Yearly_Change = Closing_Price - Opening_Price
                If Opening_Price = 0 Then
                    Percentage_Change = 0
                Else
                    Percentage_Change = Yearly_Change / Opening_Price * 100
                End If
     
             
            'Print the ticker type in the summary table
             Worksheet.Range("I" & Summary_Table_Row).Value = Ticker
     
            'Print the total volume to the Summary Table
             Worksheet.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
             
            'Print the yearly change to the Summary Table
             Worksheet.Range("J" & Summary_Table_Row).Value = Yearly_Change
                If Yearly_Change > "0" Then
                   Worksheet.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                ElseIf Yearly_Change < "0" Then
                    Worksheet.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                End If
             
               'Print the percentage change to the Summary Table
                Worksheet.Range("K" & Summary_Table_Row).Value = Percentage_Change
                Worksheet.Range("K" & Summary_Table_Row).NumberFormat = "0.00\%"
    
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
    
            'Reset the Total volume
            Total_Stock_Volume = 0
        End If
    
      Next i
  
  Next Worksheet
  
End Sub

