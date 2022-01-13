# Stock-Analysis Breakdown[Stock Analysis Report.docx](https://github.com/Jaila28/Stock-Analysis_Challenge/files/7835914/Stock.Analysis.Report.docx)


**Stock Challenge Analysis**


**Overview and Purpose of this Analysis**

One day a friend of mine name Steve reached out to me and ask for my assistance in performing an analysis.  Steve’s parents want to invest in green energy but because there are so many different companies, they need assistance. They have their eyes on one stock called DAQO. But Steve is not sure of which company would be best for them to invest in. When he told me about this idea, I thought to myself, Steve can fulfill this request himself as he just received his degree in Finance. But going through several stocks manually would be very painstaking. So, I agreed to help Steve out.

Steve needs a way to refine and narrow down his search. In this analysis I will be demonstrating to Steve how we can run scripts in excel to help him automated simple tasks. I advised to Steve that (VBA) Visual Basis Applications is a great tool to use when performing an analysis on financial data. The financial data presented will be the issue date of the stock opening, being adjusted, and closed, the stock price, ticker value, the highest and lowest price, and the volume of the stock. I will be exploring this data on Green Energy Stocks so his parents can make a proper choice on what to invest in.

Summary of the Data
The first analysis that I did for Steve. Was on the stock DAQO.

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
 'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1

 
 
When performing the analysis on DAQO. I ran a script to get the SUM of the Total Daily Volume and Return. I understood that this will help Steve’s parents get an idea of how often the DQ stock get traded.
I was able to discover that DAQO was down by 63% in 2018 which is not good.

 


Steve’s parents were highly disappointed when they found out the news. So, he provided me with two additional excel worksheets that had stock information. I was to increase the Volume of the original ticker by using Microsoft VBA to collect information on stocks from 2017 and 2018. I was able to compile data from these Green Energy Stocks through the following code below:

 
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
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
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
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
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
 
 


Steve’s parents they were excited when I advised them that I have some great news. I showed them a close comparison on the two worksheets shown below. Through this Steve parents were able to get a better picture of what stocks would best suit them for investments. This Green indicates that there was success with trades as where the Red indicates that there was a decline in trades.   

      
 
**Results of Analysis**

What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

The disadvantage of refactoring the code in this project is that it took longer than it did for me to perform this analysis I did on DAQO. There was additional code that was added and changed so that the proper outcome could manifest. But in doing so, at certain times it was difficult when figuring out how to place codes and restructuring them after adding additional strings. 
The Advantage of refactoring the code is that my code did look more easier to read and understand. This allowed the excel worksheet become neatly formatted so that Steve’s parents could understand what they were viewing, especially since the code in VBA was able to output an indication of color for the stocks that did well in trading and not so well in trading. The refactoring code also allows for debugging so that the automated tasks can run at a faster pace. 

   




