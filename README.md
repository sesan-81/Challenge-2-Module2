# Challenge2-Refactoring Code in VBA

1.	OVERVIEW OF THE ANALYSIS

Steve had previously requested that I make a workbook for his parents so they may analyze some green stocks from 2017 and 2018. The previous workbook had included the total daily volume of the stock sales as well as their average yearly returns at the click of a button. Steve loved the workbook and now wants to be able to do research on the entire stock market over the last few years.

     PURPOSE OF THE ANALYSIS
     
The purpose of this analysis is to refactor the module 2 solution code in order for it to run faster and shorter than it did. It must be made to loop through all of the stock data one time to collect information on each stock's total daily volume and the return. Moreover, the new analysis run time and evaluation must be compared to the run time of the subroutine before refactoring.

2.	RESULTS

For this refactored code, I created the ticker(tickerIndex) variable and set it equal to zero before iterating over all the rows. Using the same ticker(tickerIndex), I accessed the correct index across the four different arrays I used: the tickers array and the three output arrays which are tickerVolumes, tickerStartingPrices and tickerEndingPrices. 

The first 'for' loop was to initialize the   tickerVolumes(tickerIndex)  to zero and subsequently the loop will loop over all the rows in the spreadsheet. Inside the previous 'for' loop I wrote a script that increases the current tickerVolumes(tickerIndex) (stock ticker volume) variable and adds the ticker volume for the current volume for the current stock ticker.

 I used the ticker (Index v)ariable as the index. Using row numbering with the selected ticker(tickerIndex), I later used if-then statements to assign tickerStartingPrices and tickerEndingPrices. Using a 'for' loop through the 4 arrayes, I gave the output assignments as "Ticker", "Total Daily Volume" and "Return"
The images for 2017 and 2018 are displayed below:![image](https://user-images.githubusercontent.com/104377031/186790449-9af1d21d-d1aa-4bbb-91ec-e769729a5fc6.png)


![image](https://user-images.githubusercontent.com/104377031/186790466-90a07457-cb83-4876-b827-40cb8c75da41.png)
 







                                                                                                                                    







![Uploading image.png…]()

 
 
 
 
 

                                                                                                                                      

3.	 Summary

Based on my experience with this project, I have observed that there are both advantages and disadvantages of refactoring code in general. I found working with an original script more rewarding as it gave a 'clean slate' for creativity. Refactoring seemed to constrict the process and at times, frustrating when trying to maneauver the code or sticking to refactoring instructions, such as variable names I would have otherwise not chosen (tickerVolume vs. totalVolume).

In 2017, TERP had a negative low percentage of return as indicated by "red". 
 
In 2018 however, only ENPR and RUN had positive percentage return while others had negative percentage. almost all of the stocks had a negative percent return. 

(A)	 ADVANTAGE

•	The main advantage of refactoring would be that it gives the data analyst a template to work with.

•	It reduces execution time hugely and somewhat more efficient in nature.

(B)	DISADVANTAGE

•	The disadvantage can be seen in new challenges, such as having to recreate a button as it stopped working.

•	Complex codes used were very difficult to arrange. Also, a change to a line of code were alter the whole result and adequate care was taken to ensure no alterations were made while handling the complex codes.



SubAllStocksAnalysisRefactored()
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
    Dim tickerIndex As Integer
    tickerIndex = 0
    
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        
        tickerVolumes(tickerIndex) = 0
        
        Worksheets(yearValue).Activate
        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
                'End If
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'If  Then
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
                'End If
            End If
            
        Next i
        
    Next tickerIndex
    
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


