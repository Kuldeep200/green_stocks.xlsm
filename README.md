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
    tickerindex=0

    '1b) Create three output arrays   
    	Dim startingPrice 
   	Dim endingPrice
	Dim totalvolume
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
          If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value
            

            '3d Increase the tickerIndex. 
            tickerIndex=tickerIndex+1
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
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


Overview of Project: Explain the purpose of this analysis.
	The analysis was a request for Steve. He wants to help his parents to make the better decision to invest their money in stock.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script
![image](https://user-images.githubusercontent.com/93456209/142746510-7982eaa9-981d-431e-81c6-8f225f753864.png)

![image](https://user-images.githubusercontent.com/93456209/142746541-9bbc7f1d-9d96-435e-b090-63aa5e3cd76a.png)

	From the data ENPH (Enphase Energy Inc) had postive impact. In 2017 it had return of about 129% and in 2018 about 81.9%. This was the only stock tha t had postive impact in both years.

Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
	First Advantages is that it improve the processing data time, secondly,write fewer lines of code and create better soultions (less bugs).  The disadvantage of refactoring code is that task takes more time to think.


How do these pros and cons apply to refactoring the original VBA script?

  Pro:the processing data time was improved .The code was easy to read and understand. Con,It takes time to refactor the code, create bugs in code





