# VBA of Wall Street

## Overview of Project:
Steve’s parents are passionate about green energy, especially as fossil fuels become more scarce and energy demands shift towards renewable energy sources such as hydroelectricity, geothermal energy, wind energy, and bioenergy. They have only invested in one renewable energy company, DQ. Steve is concerned that their investments aren’t more diversified. Steve has started researching other companies in addition to DQ and put together the stock data.

Steve is interested in the total daily volume, or total number of shares traded in a day, and yearly return, or percentage difference from the beginning to year end, for the stocks he’s investigating.

First, Steve wanted to investigate the performance of DQ in 2018. Performance was measured by calculating the yearly return: percentage increase or decrease throughout the year. We found that Daqo’s stocks fell 63% in 2018, so they may not be the best option for Steve’s parents.

<img width="246" alt="Screen Shot 2022-07-26 at 8 08 40 PM" src="https://user-images.githubusercontent.com/106785377/181145387-7ec55b65-7314-4512-9e56-f0f3d8debfd6.png">

Other stocks were analyzed since Daqo didn’t seem like the best fit for Steve’s parents. The work that was used to create the DQ analysis was repurposed to analyze the other stocks for any year. Formatting was added to create an easier to read output, and a script was added to see how fast VBA could execute the code.



## Results

#### 2017 Results
<img width="225" alt="2017" src="https://user-images.githubusercontent.com/106785377/181146528-f98cf48d-9950-4dfd-920f-aaac6fc06b15.png">

#### 2018 Results
<img width="225" alt="2018" src="https://user-images.githubusercontent.com/106785377/181146540-eeeebc44-f9b8-454c-bf6c-2e14c86ea1b2.png">

The green energy stocks had better rates of return in 2017 compared to 2018 for all stocks except for ENPH and RUN. This is easy to pick out with the color indications. These indications were created using the following code.

```
dataRowStart = 4
    dataRowEnd = 15

    For x = dataRowStart To dataRowEnd
        
        If Cells(x, 3) > 0 Then
            
            Cells(x, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(x, 3).Interior.Color = vbRed
            
        End If
        
    Next x
```

Other colors can be created using [this](https://analysistabs.com/excel-vba/colorindex/) chart.

The total daily volume went up from 2017 to 2018 for some stocks but down for others. Total daily value is not the best way determine stock performance over time.

#### Run time for 2017 - Original script
<img width="248" alt="Original_Run_Time_2017" src="https://user-images.githubusercontent.com/106785377/181150346-dbbf96c2-1c3d-4d6a-bb41-8c7bbe002879.png">

#### Run time for 2018 - Original script
<img width="252" alt="Original_Run_Time_2018" src="https://user-images.githubusercontent.com/106785377/181150348-ae4bb7a0-11b9-44d7-b680-002a2f1ae602.png">

#### Run time for 2017 - Refractored script
<img width="253" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/106785377/181149460-c9ffa17c-150a-4d1b-952c-4cc66890eff1.png">

#### Run time for 2018 - Refractored script
<img width="250" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/106785377/181149472-325260e1-4ccb-4a56-a9c2-10adc3f82137.png">

The run time for stock data analyzed with the original script was 0.609375 seconds for 2017 and 0.609375 for 2018.
The run time for stock data analyzed with the refractored script was 0.5859357 seconds for 2017 and 0.59375 for 2018.

This doesn't seem like a huge difference, but this analysis was conducted for only 12 different stocks. For 100 or even 1000 different stocks, this time difference would be noticable.



## Summary


### What are the advantages or disadvantages of refactoring code?
Refractoring code is important for reducing run time and compressing Subs. Reducing run time is useful when running larger datasets. Compressing multiple Subs into one concise Sub is useful in getting the desired output without fumbling around with multiple scripts.

Although refracoring code is useful for running larger datasets, a disadvantage could be that it takes time to refine the code and consolidate it into a readable compressed form. 

### How do these pros and cons apply to refactoring the original VBA script?
The original script was simpler to create initially. It has multiple Subs, one for calculations, one for formatting, etc. It was easier to plan and parse it out this way. The refractored script was a bit more complicated to keep organized and to understand what was happening at each step in the code, and how each step would affect downstream code. Once the refractored code was complete, it was much simpler to run with just one step rather than multiple. The refractored code also ran a bit faster, which would be useful with more tickers or more stock data.
