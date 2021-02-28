# stock-analysis

# Overview of project
This project was desinged to to analyze multiple stocks to get a clearer picture of their value moving forward. the ending project was to to refactor what we had learned to make it more efficent and prove that we did. 

# Results
this is the code i used in the final challenge. 
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
' I got to this point on my own, but i couldn't figure out step 3 ( i also struggled on this part in the module as the instructions were not very clear.
'While googeling for help i came across the answer which i copied for steps 3. (https://github.com/caseychen3605/stock-analysis).
' i read over and and kinda understand the step 3 code but not percisely.i had to chane to i,8 to match the module hint.
    '3a) Increase volume for current ticker / This makes sense to me
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      '3b) Check if the current row is the first row with the selected tickerIndex.
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    '3c) check if the current row is the last row with the selected ticker/ this makes sense untill the ending price. that i dont get
       If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If
        '3d Increase the tickerIndex. assuming this is renaming to the next index?
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return. /this was copied from prior module
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i

### As noted as struglled and only was able to complete by using someone else's code for step 3. it still doesnt completey make sense to me but i do understand the end result of the code just not entirely how it got there. 
Below are the runtimes for the different years.
https://github.com/Dvallotton/stock-analysis/blob/main/Resources/2017.PNG
https://github.com/Dvallotton/stock-analysis/blob/main/Resources/2018.PNG
Overall only ENPH and RUN had a positive return in 2018, those would be the stocks to stick with. 

#### i belive refactoring is and advantage espcially when you are more confident and established with the code. You can make it easier to read and understand this way. A disadvantage is for people of my current skill can end up making it more confusing since we dont yet have a full grasp on what the code is actually doing to get the results we need. 

##### i cannot give a detailed advantage or disadvantage of the original as i do not completely understand all the code. This is what I assume a common theme among early/new developers and will only get better with more time and practice. 