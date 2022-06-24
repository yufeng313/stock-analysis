# All-Stocks-Analysis
## Overview of Project
Running an analysis of all stocks to help us better understanding "how actively tickers was traded" and "how different tickers performed" and in last few years. In this challenge, I'll use VBA code to automate tasks in a faster way.
## Results
In order to analysis all the dataset, here are some purpose we need to bring in mind:<br/>
1.Create arrays for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices by using tickerIndex.<br/>
2.Use for loop to caculate the total volume and year return for every tickers.<br/>
3.Use Timer function to know how fast VBA code will compile the results.
And here are the results with some of the VBA code:
![main code of stocks analysis](https://user-images.githubusercontent.com/107179765/175507746-e62bd7da-43ec-4c8e-ac98-f28f1be09342.png)
![Stock Analysis Output for 2017](https://user-images.githubusercontent.com/107179765/175507780-7a335063-9c33-4a7f-8fe2-1184dbd3e6e1.png)<br/>
![the running time of original script 2017](https://user-images.githubusercontent.com/107179765/175652522-09f3f572-effc-478c-9250-dc651854925f.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107179765/175507826-dba2ae97-b603-422e-a64f-11f6d9175d8d.png)<br/>
![Stock Analysis Output for 2018](https://user-images.githubusercontent.com/107179765/175507805-93195e55-53b3-465b-b978-4658704b996d.png)<br/>
![the running time of original script 2018](https://user-images.githubusercontent.com/107179765/175652586-323694ab-c44e-4a5e-bc1a-32dd2d68f483.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107179765/175507846-59cc0c6e-1e46-4129-9d9c-beef4eefea6c.png)
### Analysis Results
From the output results, we can see that in year 2017,all of the tickers have a positive return except for the ticker "TERP" has a negative return of -7.2%,but in year 2018,most company had a negative returns, which means all stocks had a better performance at year 2017.<br/> 
And we can also see that the ticker "ENPH" had great total volumes and returns in both year. This concludes with the suggestion to purchase stocks from "ENPH" maybe a good choice.<br/>
After running the stocks analysis code ,the execution time of the refactored script is less than the original script, since the refactored script only takes 0.15 seconds.
## Summary
-What are the advantages or disadvantages of refactoring code?<br/>
 The advantage of refactoring code is that you can make an analysis of any year without having to write a code for every year, and this really save the coding time and performing time on it.<br/>
 The disadvantage of refactoring code is that the code become more complex, and it might be confused the next time when we exposed to the code.<br/>
-How do these pros and cons apply to refactoring the original VBA script?<br/>
 We can use nested loop and create index or variables to make the code looks shorter and more precise,this will reduce the memory usage and make the code running faster.But this also means that the refactored script requires more complex thinking of dealing with different variables and functions, which will increase the thinking process of putting all the data together in a correct order,thus we may use pseudocode to clear our mind before writing our code.
