# Stock-Analysis
**Project Overview**

Steve asked me to perform a green stock analysis so that he can let his parents know if it is worth investing in. To do this I used the Visual Basic Application or VBA in excel to create macros, which are lines of code that are executed within the document. I used VBA to find the stock’s total daily volume and annual return. Then I added to the code to analyze 11 more stocks to see how the one in question compared and this allowed Steve to determine if it was worth investing in. 

**Purpose**
The purpose of this project was to take the VBA script created in the module and make it more efficient. After creating the entire VBA script to analyze the 12 stocks it was determined that there was a more efficient way of running the code. Taking an existing code and adding or subtracting elements is called refactoring. By refactoring the code, I was able to make it more efficient and have it run in less time. 

**Results**
To make the code run more efficient first I created a tickerIndex variable and set that equal to 0. Then I created three output arrays; tickerVolumes, tickerStartingPrices, and tickerEndingPrices. This was to establish the ticker symbol of stock. By having the tickerIndex variable I was able to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before iterating through the data set rather than having to use a nested loop for each. 

**Oringial Code**

![AllStock1](https://user-images.githubusercontent.com/118235205/204983088-0aae0934-27b4-4a38-aad4-3ed65e75594f.png)
![AllStock2](https://user-images.githubusercontent.com/118235205/204983103-28d961e6-07f1-42ca-92d5-97f4602ef39c.png)
![AllStock3](https://user-images.githubusercontent.com/118235205/204983121-19e1e2fb-4979-4015-8df2-a9853c8f8f3c.png)

**Refactored sections of code**

![Refrac1](https://user-images.githubusercontent.com/118235205/204983196-0b4c0944-7058-4b6e-af20-59a5bed36da3.png)
![Refrac2](https://user-images.githubusercontent.com/118235205/204983218-ece87f2f-7248-4d6d-9fc1-195b36c9baf9.png)

**Run-time for Original and Refactored Code**

![2017_Original](https://user-images.githubusercontent.com/118235205/204983268-c08338a2-cc71-4c13-8292-216358f252b8.png)

![2017_Refactored](https://user-images.githubusercontent.com/118235205/204983255-20740995-6281-4090-87d9-58aecb41dc3f.png)

<img width="384" alt="2018_Original" src="https://user-images.githubusercontent.com/118235205/204983335-3680a622-7244-4937-a194-8580ac34fa69.png">

<img width="482" alt="2018_Refactored" src="https://user-images.githubusercontent.com/118235205/204983356-43b61b52-2317-465d-ae7b-db860712ee80.png">


When comparing the original code and the refactored code I saw a speed increase for both 2017 and 2018 data. The original code ran in 0.578125 seconds for 2017, and the refactored ran in 0.0625 seconds. For 2018, the original code ran in 0.546875 seconds and the refactored in 0.0703125 seconds. The refactored code saved about 0.5 seconds for each year, thus making it more efficient. 
     
**Summary**

**Refactoring Advantages/Disadvantages In general**

Refactoring has a major advantage of making code more efficient and taking out redundancies. This is helpful when running large codes that take hours. The major disadvantage is that you are taking an already working code and changing it, potentially making it unusable. It is for this reason to always save copies of the original working code in case it cannot be refactored correctly. 

**Refactoring Advantages/Disadvantages In VBA**

A major advantage to refactoring code in VBA is that it is easy to create a new module and add the original code to it and put in the new additions and compare the new and old side by side to find potential areas of improvement. A disadvantage is that if the code is already efficient, refactoring may only shave off seconds, requires a lot of work, and potentially can render the code useless. Additionally if you are working on someone else’s code and do not understand the entire syntax and layout you may make it less efficient when trying to do the opposite. 



