Attribute VB_Name = "Module1"


'Sub Function for Stocks Data to Work on Multi Year Data to Extract Different Summary Calucations and Formatting Table


Sub Stocks():

    For Each ws In Worksheets
    
    
    'Declaring variables here
    
    Dim TickerSymbol As String
    
    Dim YearlyChange As Double
    
    Dim PercentChange As Double
    
    Dim TotalStock As Double
    
    Dim GreatestIncrease As Double
    
    Dim GreatestDecrease As Double
    
    Dim GreatestTotalVolume As Double
    
    Dim OpenValue As Double
    
    Dim OutputTable As Integer

     
    'Get the worksheets to used al at once
     
     WorksheetName = ws.Name

    
    
    'Initializing the varibales
    
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestTotalVolume = 0
    
    
    'Set the Output Table Value
    
    OutputTable = 2
    
    OpeningPrice = ws.Cells(2, 3).Value
    
    
    ' Determine the Last Row
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Header for the OutputTable Columns
    
      ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 11).Value = "Percent Change"
      ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'Column Headers for Summary Table Calucation
     
     ws.Cells(1, 15).Value = "Summary"
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"
    
    
    'Row Title for Summary Table Calculations
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    
    'Set the total stock initial value
    
    TotalStock = 0

    
    OpenValue = ws.Cells(2, 3).Value
    
    
    'Looping through second to last row


    For i = 2 To LastRow
    
    TotalStock = TotalStock + ws.Cells(i, 7).Value
    
    
    'Checking for change in tikcer values
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
         ws.Cells(OutputTable, 9).Value = ws.Cells(i, 1).Value
         
         ws.Cells(OutputTable, 12).Value = TotalStock
         
         
         'Calculate and show the Yearly Change in Output Table
         
         YearlyChange = ws.Cells(i, 6).Value - OpenValue
         
         ws.Cells(OutputTable, 10).Value = YearlyChange
         
         'Format the Yearly Change with $ symbol
         
         ws.Cells(OutputTable, 10).NumberFormat = "$#,##0.00"
         
         'Conditional Formatting for Yearly Change Column
         
         Select Case YearlyChange
         
         Case Is >= 0
            ws.Cells(OutputTable, 10).Interior.ColorIndex = 4
         
         Case Is < 0
            ws.Cells(OutputTable, 10).Interior.ColorIndex = 3
         
         Case Else
            ws.Cells(OutputTable, 10).Interior.ColorIndex = 0
         
         End Select
         
            
            
            
            
         'Calculate and show the Percent Change in Output Table
         
            PercentChange = YearlyChange / OpenValue
         
            ws.Cells(OutputTable, 11).Value = PercentChange
          
         'Format the Percent Change Columns in Percentage
          
            ws.Cells(OutputTable, 11).NumberFormat = "0.00%"
         
         'Conditional Formatting for Percent Change Column
         
         'Select Case PercentChange
         
         'Case Is >= 0
          '   ws.Cells(OutputTable, 11).Interior.ColorIndex = 4
         '
         'Case Is < 0
          '   ws.Cells(OutputTable, 11).Interior.ColorIndex = 3
         
         'Case Else
          '   ws.Cells(OutputTable, 11).Interior.ColorIndex = 0
         '
         'End Select
         
         
         
        '____Summarize the table required values______
        
        'To find out the Greatest Percent Increase
    
        If ws.Cells(OutputTable, 11).Value > GreatestIncrease Then
    
            ws.Cells(2, 17).Value = ws.Cells(OutputTable, 11).Value
    
            GreatestIncrease = ws.Cells(OutputTable, 11).Value
    
            ws.Cells(2, 16).Value = ws.Cells(OutputTable, 9).Value
            
            'Format the value to percentage
            
                ws.Cells(2, 17).NumberFormat = "0.00%"
        
        End If
    


        'To find out the Greatest Percent Decrease
    
        If ws.Cells(OutputTable, 11).Value < GreatestDecrease Then
            
            ws.Cells(3, 17).Value = ws.Cells(OutputTable, 11).Value
    
                GreatestDecrease = ws.Cells(OutputTable, 11).Value
    
            ws.Cells(3, 16).Value = ws.Cells(OutputTable, 9).Value
            
            'Format the value to percentage
                ws.Cells(3, 17).NumberFormat = "0.00%"
    
        End If
    
          
         
        'To find out the Greatest Total Volume

         If ws.Cells(OutputTable, 12).Value > GreatestTotalVolume Then
    
            ws.Cells(4, 17).Value = ws.Cells(OutputTable, 12).Value
    
                GreatestTotalVolume = ws.Cells(OutputTable, 12).Value
    
            ws.Cells(4, 16).Value = ws.Cells(OutputTable, 9).Value

        End If
         
           'Format the Greatest Total Volume
                ws.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
          
         
         ' Change values to prepare for next ticker
         
         OutputTable = OutputTable + 1
         
         TotalStock = 0
         
         OpenValue = ws.Cells(i + 1, 3).Value
    
            
            
    End If

    Next i
    
        
        'Format columns to auto adjust the

        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
Next ws
End Sub

