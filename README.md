# stock-analysis

Overview of Project

A friend name Steve became a financial planner and decided to help his parents analyze some stock for their portfolio. His parents favored the stock with the ticker DQ due to personal reasons. Steve wanted to present data on the performance of the DQ stock to his parents. Unfortunately the stock did not perform well. Because of this, Steve asked for an analysis of 12 different stocks.

The stock analysis project was created to compare the total daily volume and the return of 12 different stocks. Data was taken from the performance of various stocks throughout 2017 and 2018. This included over 3000 lines of data. Code was written in VBA script to allow Steve to run a report of stocks for both 2017 and 2018. A button was added so Steve could run his reports with little effort. This original report is entitiled "Green Stocks."

The speed at which the data was delivered was important to Steve. To help optimize the speed the original code used was refractored to become more efficient for Steve's purposes. The new faster report is called "VBA CHallenge." Refractoring code is the act of reusing previously written code while making changes to apply to a new project.


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
         'If the next row???s ticker doesn???t match, increase the tickerIndex.
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
![image](https://user-images.githubusercontent.com/105091538/172523675-f8e1bbf9-f8ea-44ae-b0c3-03cf1385de2a.png)


Green Stocks 2018
![image](https://user-images.githubusercontent.com/105091538/172523791-f5ed3dbc-d3f7-42a9-bc05-2f9080075513.png)


VBA Challenge 2017
![image](https://user-images.githubusercontent.com/105091538/172523840-06e83395-2b99-4051-ac4a-ae6e24e37a6c.png)


VBA Challenge 2018
![image](https://user-images.githubusercontent.com/105091538/172523893-584ba700-fd69-4c51-bce1-5872bbee9305.png)



Summary

Advantages and Disadvantages of Refractoring Code

This project was completed using refractored code. There are both advantages and disadvantages to refractoring code. Using code that has already been written saves time. This means that most of the code is already there and only a few changes are needed. This also helps prevent errors. It is a lot easier to make a coding mistake when rewriting an entire piece of code.

A disadvantage of refractoring code is that sometimes just changing pieces of code does not work. This can lead to debugging the code to find the issue. It is important to test the code as you go to ensure there are no mistakes.


Advantages and Disadvantages 

There are both advantages and disadvantages of the original and refractored VBA script. The original VBA script runs well. It provides the information requested without issue. The disadvantage is the time to run the report. The time it takes to run the report is around 0.27 seconds. That is not a long time, but the time could definitely be shortened.

An advantage to using the refractored VBA script is that only takes 0.07 seconds to run. That's a big time savings. It also was easy to write since it used a lot of original code. A disadvantage is that the refractored VBA script may not produce as nice a report as the original script. 

