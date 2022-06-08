# stock-analysis

Overview of Project

A friend name Steve became a financial planner and decided to help his parents analyze some stock for their portfolio. His parents favored the stock with the ticker DQ due to personal reasons. Steve wanted to present data on the performance of the DQ stock to his parents. Unfortunately the stock did not perform well. Because of this, Steve asked for an analysis of 12 different stocks.

The stock analysis project was created to compare the total daily volume and the return of 12 different stocks. Data was taken from the performance of various stocks throughout 2017 and 2018. This included over 3000 lines of data. Code was written in VBA script to allow Steve to run a report of stocks for both 2017 and 2018. A button was added so Steve could run his reports with little effort. This original report is entitiled "Green Stocks."

The speed at which the data was delivered was important to Steve. To help optimize the speed the original code used was refractored to become more efficient for Steve's purposes. The new faster report is called "VBA CHallenge."


Results


Original Code

The original code used to populate the "Green Stocks" report is below. This code ran the 2017 report in 0.265625 seconds. The 2018 report ran in the same length of time. 


Sub AllStocksAnalysis()

   '1) Format the output sheet on All Stocks Analysis worksheet
   
   Worksheets("All Stocks Analysis").Activate
   
        Dim startTime As Single
        
        Dim endTime As Single
   
    yearValue = InputBox("What year would you like run the analysis on?")
    
        startTime = Timer
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   
   'Create a header row
   
   Cells(3, 1).Value = "Ticker"
   
   Cells(3, 2).Value = "Total Daily Volume"
   
   Cells(3, 3).Value = "Return"

   
   '2) Initialize array of all tickers
   
   Dim tickers(11) As String
   
   tickers(0) = "AY"
   
   tickers(1) = "CSIQ"
   
   tickers(2) = "DQ"
   
   tickers(3) = "ENPH"
   
   tickers(4) = "FSLR"
   
   tickers(5) = "HASI"
   
   tickers(6) = "JKS"
   
   tickers(7) = "RUN"
   
   tickers(8) = "SEDG"
   
   tickers(9) = "SPWR"
   
   tickers(10) = "TERP"
   
   tickers(11) = "VSLR"
   
   
   '3a) Initialize variables for starting price and ending price
   
   Dim startingPrice As Single
   
   Dim endingPrice As Single

   
   '3b) Activate data worksheet
   
   Worksheets(yearValue).Activate

   
   '3c) Get the number of rows to loop over
   
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   
   '4) Loop through tickers
   
   For i = 0 To 11
       
       ticker = tickers(i)
       
       totalVolume = 0
       
       
       '5) loop through rows in the data
       
       Worksheets("2018").Activate
       
       For j = 2 To RowCount
           
           '5a) Get total volume for current ticker
           
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           
           '5b) get starting price for current ticker
           
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       
       Next j
       
       
       '6) Output data for current ticker
       
       Worksheets("All Stocks Analysis").Activate
       
       Cells(4 + i, 1).Value = ticker
       
       Cells(4 + i, 2).Value = totalVolume
       
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    

End Sub




Sub formatAllStocksAnalysisTable()

    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#, ##0"
    
    Range("C4:C15").NumberFormat = "$0.00%"
    
    Columns("B").AutoFit


    dataRowStart = 4
    
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            'Color the cell green
            
            Cells(i, 3).Interior.Color = vbGreen
            
            
        ElseIf Cells(i, 3) < 0 Then
        
        
            'Color the cell red
            
            Cells(i, 3).Interior.Color = vbRed
            
        Else
        
            'Clear the cell color
            
            Cells(i, 3).Interior.Color = xlNone
            
        End If
        
    Next i
      
End Sub



Updated Code

The refractored code used to populate the "VBA Challenge" report is below. The new reports for both 2017 and 2018 run in 0.0703125 seconds.



Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     
         End If

            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                 tickerIndex = tickerIndex + 1
                 
        End If

        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Images of Time Elapsed

Green Stocks 2017
Resources/Resources/Green_Stocks_2017_Run_Time.png 

Green Stocks 2018

VBA Challenge 2017

VBA Challenge 2018


The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
