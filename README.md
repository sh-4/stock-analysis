# **VBA Stock Analysis**

## **Overview of Project**
The purpose of this analysis is to quickly access information about stock market data that is forever changing. By creating a fluid code, someone could quickly compare how these varying stocks perform throughout the years, without needing to have any knowledge of coding to adjust for new spreadsheets. In the stock market game, time is money, so creating a code that can sort through thousands of cells of data in a tenth of a second will help to meet the demands and expectations within this field.
### **Results** 
Using the code created, it’s easy to compare the data of multiple stocks between 2017 and 2018. 

![StockAnalysis_2017](https://user-images.githubusercontent.com/105808695/174675740-7ff1aaaa-bfd4-48a0-abf2-7181ed1655ca.png)

As you can see, 2017 was a good year for most of the companies, with TERP being the only one producing a negative return. Even though some of the companies had a smaller volume, or traded shares, the return could still come back as a growth or decline in the investment.

![StockAnalysis_2018](https://user-images.githubusercontent.com/105808695/174675775-237a7ba9-4a39-490b-801b-cae58551350c.png)

Looking at the data for the following year, 2018, we can determine that having a higher daily volume does not correlate to whether or not the return will be a positive investment. DQ’s daily volume tripled between 2017 and 2018, yet the return reflected a significant decrease. The same holds true for HASI, SEDG, TERP, and VSLR increasing in daily volume, yet decreasing in their return. On the opposite side, ENPH and RUN both had increases in their volume, but remained positive in their return, even though it was a smaller increase than they had in 2017.

Collecting this data allows the investor to watch multiple companies and track their progress, or lack of, through the years. Currently, it shows that ENPH and RUN seem to be better companies to invest in, whereas DQ, JKS, and SPWR are trending downward. TERP may be worth keeping an eye on in the future, as it's return actually did trend upwards slightly, but has not yet shown promising growth over this two-year timespan.

This information was originally coded to gather information on these specific companies by going through all of the data 12 different times. While it was under 1 second per year analyzed, as seen below, it is still important to remember that time is money in the stock exchange and there is likely a more efficient way to get the same data.

![GreenStocks_2017](https://user-images.githubusercontent.com/105808695/174675820-a6c4478f-a3aa-45c8-a17b-b48ccc41425f.png)
![GreenStocks_2018](https://user-images.githubusercontent.com/105808695/174675835-7c380e28-678c-4469-9940-8f756e48bb99.png)

By refactoring, or editing, the original code, the exact same information is able to be retrieved 12.5% faster. Now instead of waiting 0.8 seconds, the user will only have to wait approximately 0.1 seconds.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/105808695/174675879-5f5f773a-1be8-4b7a-9f60-511418fac34c.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/105808695/174675887-71a6b476-f38a-4f6d-affc-6a41d3a125c4.png)

All it took to get these faster results was a slight adjustment in the coding. Rather than looping through thousands of rows of data multiple times, it’s now refactored to get all of the same information in one swift run. As years pass, more information will likely be added, so it’s important to have the code set up in anticipation of the sheets growing as time moves forward. If not, it’s possible that by 2025 it would take 5 to 10 seconds to loop through all of the data, which would be unacceptable to those wanting access to the information as quickly as possible.
## **Summary**
### **Refactoring Code: Advantages and Disadvantages**
Refactoring code gives the advantage of fine-tuning a first draft. It’s important to ensure the code is working correctly and accurately before going back through to see what efficiencies can be made. Creating an original code also allows the coder to see what the end result needs to look like, so they can make sure their refactored version matches. When a code is refactored well, it can save the coder and end-user from wasting unnecessary time awaiting results; this is especially important as the data sets get larger. A disadvantage of refactoring is the new possibility that something integral to the structure of the code could be altered, thus rendering the code broken or the information incorrectly translated.
### **How the VBA Script was Affected by the Pros and Cons**
The original VBA script took longer to loop through the information to give its results. By refactoring, the biggest advantage was able to be executed: more efficiency. The time it took to get the information transitioned from 0.8 seconds down to 0.1 seconds. Additionally, the cons also made an appearance before the final result. I had a few moments where the code was altered incorrectly, causing it to not run through at one point, and to not give the correct information at another. Luckily, the benefit of having access to the original code was that I could compare my results from my refactored code until the data matched correctly and it ran as intended.

