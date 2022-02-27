# Stock-Analysis
Using Visual Basic for Applications (VBA) to perform detailed analysis on green energy stocks, such as "DQ." The importance on our reliance on alternative energy sources is rapidly increasing as we consume more fossil fuels.  Our analysis specifically focuses on green energy stocks in 2017 and 2018, unconvering their yearly return and volume in order to dictate which stocks would be best to currently invest in.  For easier visualization, we highlighted those tickers with a negative return in red and those with a positive return in green.
# Overview of Project
The purpose of this project was to learn how to refractor an original script to perform multiple tasks with one function.  We were able to refractor our original script to be able compare stock performance as well as determine code execution times when code is ran.  Through this process, we will be able to learn if refractoring a code is more beneficial in terms of improving the execution times for the VBA script.
# Results
To view the more detailed, raw data showing the perofrmance of the stocks in 2017 and 2018, you may click [here](https://github.com/nataliabench/Stock-Analysis/blob/2c1518802c039d1675b4182ad4e5c90cfa346b83/VBA_Challenge.xlsm).

After running our yearly analysis code for all stocks in 2017 and 2018, we recieved vastly differing results in terms of negative and positive turns.  In the image below to the left, we can see a table showing a list of stock tickers, total daily volume, and return.  This image shows that out of the 12 green energy stocks analyzed in 2017, only one had a negative return.  The "Return" column is indiciating the change in price from the start of the year to the end.  As for the table below on the right, we can infer that in 2018, almost all the green energy stocks we analyzed gave us a negative return for the year.  Only two green energy stocks had a positive return around 80%.  We were able to create a more readable and visualized interpretation ofthe data through conditional formatting in VBA.  We created a script that instructed the computer to highliter a cell red if itwas below zero and to highlight the cell green if it was above zero. Therefore, when comparing both tables, we can conclude that "RUN" and "ENPH" would be the most reliable stocks to invest in as they were the only ones that both maintained positive returns in 2017 and 2018.
![alt text](https://github.com/nataliabench/Stock-Analysis/blob/6f3f8856e85ebb2017d9c8f91858ed20cbc5a0f1/Resources/2017:2018%20All%20Stock%20Analysis.png)
Our original script contained a nested for loop that was able to perform and repeat the lines of code until told to stop when stated in code.  We decided to refractor our original code, in hopes of improving the time elapsed for code run time when creating pop-up message boxes.  We wanted to improve our code so that we would recieve the data on the run time quicker.  To do so, we refracted our code to instead be scripted to loop through all the data at once to perform the same function as originally.  
Once we ran our refracted script, we saved our pop-up message boxes to remember the run time for year 2017 and 2018.  Then, I ran our original script again, and saved the run times as well for both years.  Below, you can see the run times with our refracted code for years 2017 and 2018.

![alt text](https://github.com/nataliabench/Stock-Analysis/blob/c9daefed62704fceb11da3e89899f129e1f07053/Resources/VBA_Challenge_2017.png)
![alt text](https://github.com/nataliabench/Stock-Analysis/blob/c9daefed62704fceb11da3e89899f129e1f07053/Resources/VBA_Challenge_2018.png)

The difference between the run times for the original script vs. refracted script were minimal but, still we saw a slower run time with our refracted code for both years. Below, you can find the original code run times.

![alt text](https://github.com/nataliabench/Stock-Analysis/blob/590508a400b67972e389e55002c23652f5abbc7b/Resources/Original_Script_RunTime2017.png)
![alt text](https://github.com/nataliabench/Stock-Analysis/blob/590508a400b67972e389e55002c23652f5abbc7b/Resources/Original_Script_RunTime2018.png)


# Summary
Disadvantages and advantages of refractoring code
Disadvantages and advantages on the original code vs. refractored code
