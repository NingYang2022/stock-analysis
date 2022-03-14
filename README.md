# Refactor VBA Code and Measure Performance

## <font color=#6495ED>Overview of Project</font>
Steve’s parents like green energy very much. They think that alternative energy productions are promising as fossil fuels get used up. They decided to invest in some green companies. However, they haven’t done much research and decided to invested all their money in one company DQ, Steve'd like to diversify their funds, he wants to analyze a few green energy stocks. He created an excel file containing the stock data and asked for help with VBA code to automate analyses.
### <font color=#6495D>Purpose</font>
* To find out which companies are worth investing based on comparision of the stock performance between 2017 and 2018 
* To find out code performance improvement based on comparision of execution time between original and refactored scripts.

---
## <font color=#6495ED>Results</font>

1. Only two green companies had positive returns on both years, ENPH and RUN, which show good performance and will be recommended to be invested in.
![Stock_performance_2017](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/Stock_performance_2017.png?raw=true)
![Stock_performance_2018](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/Stock_performance_2018.png?raw=true)

2. The refactored script runs faster than the original one.
- Images of pop message box for 2017 and 2018 of the original code show that the script used more time to run. 
![VBA_Original_2017](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/VBA_Original_2017.png?raw=true)
![VBA_Original_2018](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/VBA_Original_2018.png?raw=true)
- Images of pop message box for 2017 and 2018 of the refactored code, on the other hand, show that the script used less time to run.
![VBA_Challenge_2017](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)
![VBA_Challenge_2018](https://github.com/NingYang2022/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

3. The refactored script manipulates the  added three arrays with “tickerIndex” and removed the related outer ‘For’ loop from the original code, making the execution time less
```
 '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
```
```
For i = 2 To RowCount
            
    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i - 1, 1) <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
            
    '3c) check if the current row is the last row with the selected ticker
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
    If Cells(i + 1, 1) <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        '3d) increase the tickerIndex
        tickerIndex = tickerIndex + 1
                    
    End If
        
Next i
```
## <font color=#6495ED>Summary</font>

### <font color=#6495ED>1. Pros and Cons of refactoring code in general</font>
- Advantages:
    -	Make code run faster.
    -	Make code clean and organized, easier to read, understand, debug and maintain.
    -	Cost less system resources.


- Disadvantages:
    -	It’s risky when the developer does not understand well about the project
    -	Software is hard to refactor when it is big. Refactoring will cost time and money.

### <font color=#6495ED>2. Pros and Cons of refactoring the original VBA script in this project</font>
- Advantages:
    -	The refactored script runs faster.
    -	The refactored script  is organized  and broken down into  smaller units, which is easier to maintain
    

- Disadvantages:
    - Developers have to understand  what the project is  all about very well before refactoring
    - It cost time when re-design the software by adding arrays and removing the outer "For" loop. It needs to pay more attention to the arrays and index in them.
